package main

import (
	"context"
	"errors"
	"fmt"
	"runtime"
	"sync"
	"time"

	ole "github.com/go-ole/go-ole"
)

// Alle COM-Aufrufe laufen auf einem einzigen OS-Thread im Single-Threaded
// Apartment (STA). Die Tool-Handler schicken
// ihre Requests ueber einen Kanal dorthin; damit sind sie auch serialisiert —
// Outlooks COM-Schnittstelle ist ohnehin single-threaded.
//
// Lebensdauer: Der Thread startet lazy beim ersten Tool-Call. Nach idleTimeout
// ohne Call gibt er die Outlook-Referenzen frei, damit outlook.exe nicht als
// unsichtbarer Geisterprozess weiterlaeuft, nachdem der Nutzer das Fenster
// geschlossen hat. Haengt Outlook laenger als callTimeout (z. B. an einem
// Dialog), wird der Thread aufgegeben und der naechste Call startet einen
// neuen — der alte raeumt auf, sobald Outlook ihn wieder loslaesst.
const (
	idleTimeout = 5 * time.Minute
	callTimeout = 90 * time.Second
	maxTries    = 6
)

type staJob struct {
	op     string
	params map[string]any
	done   chan staResult // Puffer 1: ein aufgegebener Thread blockiert nie
}

type staResult struct {
	out []byte
	err error
}

type staThread struct {
	jobs   chan staJob
	quit   chan struct{}
	exited chan struct{}
}

type outlookCOM struct {
	mu   sync.Mutex
	th   *staThread
	dial func() (backend, error)
	// sleep ist im Test austauschbar, damit Retries nicht Sekunden kosten.
	sleep func(time.Duration)
	dir   dirCache
}

var outlook = &outlookCOM{dial: dialOutlook, sleep: time.Sleep}

// call fuehrt eine Op aus und liefert ihre JSON-Lines-Ausgabe.
func (c *outlookCOM) call(ctx context.Context, op string, params map[string]any) ([]byte, error) {
	c.mu.Lock()
	th := c.th
	if th == nil {
		th = &staThread{jobs: make(chan staJob), quit: make(chan struct{}), exited: make(chan struct{})}
		c.th = th
		go c.loop(th)
	}
	c.mu.Unlock()

	timer := time.NewTimer(callTimeout + syncGrace(op, params))
	defer timer.Stop()
	timeout := fmt.Errorf("timeout in '%s' — Outlook does not respond "+
		"(possibly stuck on a dialog). Connection discarded", op)

	j := staJob{op, params, make(chan staResult, 1)}
	select {
	case th.jobs <- j:
	case <-ctx.Done():
		return nil, ctx.Err()
	case <-timer.C:
		c.abandon(th)
		return nil, timeout
	}

	select {
	case r := <-j.done:
		if r.err != nil {
			return nil, fmt.Errorf("Outlook/COM error: %s", r.err)
		}
		return r.out, nil
	case <-ctx.Done():
		// Die Op laeuft auf dem STA-Thread zu Ende; der Thread bleibt brauchbar.
		return nil, ctx.Err()
	case <-timer.C:
		c.abandon(th)
		return nil, timeout
	}
}

// syncGrace: mail_sync wartet absichtlich, das darf nicht als
// Zeitueberschreitung durchgehen.
func syncGrace(op string, params map[string]any) time.Duration {
	if op != "sync" {
		return 0
	}
	if wait, ok := params["wait"].(int); ok && wait > 0 {
		return time.Duration(wait) * time.Second
	}
	return 0
}

func (c *outlookCOM) abandon(th *staThread) {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.th == th {
		close(th.quit)
		c.th = nil
	}
}

// stop gibt Outlook beim Herunterfahren frei und wartet kurz darauf.
func (c *outlookCOM) stop() {
	c.mu.Lock()
	th := c.th
	c.th = nil
	c.mu.Unlock()
	if th == nil {
		return
	}
	close(th.quit)
	select {
	case <-th.exited:
	case <-time.After(3 * time.Second):
	}
}

func (c *outlookCOM) loop(th *staThread) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	defer close(th.exited)

	initErr := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	var oe *ole.OleError
	if errors.As(initErr, &oe) && oe.Code() == 1 { // S_FALSE: war schon initialisiert
		initErr = nil
	}
	if initErr == nil {
		defer ole.CoUninitialize()
	}

	var be backend
	drop := func() {
		if be != nil {
			be.close()
			be = nil
		}
	}
	defer drop()

	idle := time.NewTimer(idleTimeout)
	defer idle.Stop()
	for {
		select {
		case j := <-th.jobs:
			if initErr != nil {
				j.done <- staResult{nil, fmt.Errorf("COM initialization failed: %w", initErr)}
				continue
			}
			out, err := c.run(&be, j.op, j.params)
			j.done <- staResult{out, err}
			idle.Reset(idleTimeout)
		case <-idle.C:
			drop()
		case <-th.quit:
			return
		}
	}
}

// run fuehrt eine Op aus und wiederholt sie bei transienten COM-Fehlern mit
// frischer Verbindung. Die Ausgabe eines Versuchs zaehlt nur, wenn er gelingt —
// sonst kaemen Zeilen nach einem Retry doppelt.
func (c *outlookCOM) run(be *backend, op string, params map[string]any) ([]byte, error) {
	for try := 1; ; try++ {
		out, err := c.attempt(be, op, params)
		if err == nil {
			return out, nil
		}
		if !isTransient(err) || try >= maxTries {
			return nil, err
		}
		if *be != nil {
			(*be).close()
			*be = nil
		}
		c.sleep(time.Duration(400*try) * time.Millisecond)
	}
}

func (c *outlookCOM) attempt(be *backend, op string, params map[string]any) ([]byte, error) {
	if *be == nil {
		b, err := c.dial()
		if err != nil {
			return nil, err
		}
		*be = b
	}
	ol, ns := (*be).begin()
	defer (*be).end()
	s := &session{ol: ol, ns: ns, dir: &c.dir}
	if err := s.do(op, params); err != nil {
		return nil, err
	}
	return s.out.Bytes(), nil
}
