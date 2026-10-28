package main

import (
	"errors"
	"fmt"
	"math"
	"strings"
	"syscall"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// obj ist ein Outlook-Objekt, so wie die Ops es sehen. Die echte Fassung ist
// comObj (IDispatch ueber go-ole), in den Tests eine Faelschung (fake_test.go).
// Rueckgabewerte sind string, int, float64, bool, time.Time, obj oder nil.
type obj interface {
	get(name string, args ...any) (any, error)
	call(name string, args ...any) (any, error)
	put(name string, v any) error
}

// backend ist eine Verbindung zu Outlook. begin liefert Application und
// Namespace fuer einen Request, end gibt alles frei, was der Request unterwegs
// geholt hat, close die Verbindung selbst.
type backend interface {
	begin() (ol, ns obj)
	end()
	close()
}

// arena sammelt die IDispatch-Referenzen eines Requests. Freigegeben wird am
// Ende des Requests in einem Rutsch — so kann keine Op eine Referenz vergessen,
// und outlook.exe bleibt nicht wegen einer liegengebliebenen Referenz am Leben.
type arena struct{ refs []*ole.IDispatch }

func (a *arena) release() {
	for i := len(a.refs) - 1; i >= 0; i-- {
		a.refs[i].Release()
	}
	a.refs = nil
}

type comObj struct {
	d *ole.IDispatch
	a *arena
}

func (o *comObj) get(name string, args ...any) (any, error) {
	v, err := oleutil.GetProperty(o.d, name, comArgs(args)...)
	return o.wrap(v, err)
}

func (o *comObj) call(name string, args ...any) (any, error) {
	v, err := oleutil.CallMethod(o.d, name, comArgs(args)...)
	return o.wrap(v, err)
}

func (o *comObj) put(name string, v any) error {
	r, err := oleutil.PutProperty(o.d, name, comArgs([]any{v})...)
	if r != nil {
		r.Clear()
	}
	return comErr(err)
}

// wrap macht aus einem VARIANT einen Go-Wert. Ein Objekt behaelt seine
// Referenz und landet in der Arena; alles andere wird kopiert und der VARIANT
// sofort freigegeben (BSTR).
func (o *comObj) wrap(v *ole.VARIANT, err error) (any, error) {
	if err != nil {
		return nil, comErr(err)
	}
	if v == nil {
		return nil, nil
	}
	if v.VT == ole.VT_DISPATCH {
		d := v.ToIDispatch()
		if d == nil {
			return nil, nil
		}
		o.a.refs = append(o.a.refs, d)
		return &comObj{d, o.a}, nil
	}
	if v.VT == ole.VT_DATE {
		return fromOADate(math.Float64frombits(uint64(v.Val))), nil
	}
	val := v.Value()
	v.Clear()
	switch x := val.(type) {
	case int8:
		return int(x), nil
	case uint8:
		return int(x), nil
	case int16:
		return int(x), nil
	case uint16:
		return int(x), nil
	case int32:
		return int(x), nil
	case uint32:
		return int(x), nil
	case int64:
		return int(x), nil
	case uint64:
		return int(x), nil
	case float32:
		return float64(x), nil
	case *ole.IUnknown:
		return nil, nil // schon per Clear freigegeben, nichts Brauchbares
	}
	return val, nil
}

// fromOADate rechnet ein OLE-Datum (Tage seit 30.12.1899, Bruchteil = Uhrzeit)
// wie .NET DateTime.FromOADate auf Millisekunden genau um. go-oles eigene
// Umrechnung (VariantTimeToSystemTime) rundet auf ganze Sekunden — 14:13:59,6
// wuerde so zu 14:14.
func fromOADate(d float64) time.Time {
	const msPerDay = 86400000
	ms := int64(d*msPerDay + 0.5)
	if d < 0 {
		ms = int64(d*msPerDay - 0.5)
		// Vor 1899: der Bruchteil zaehlt vom Tag aus vorwaerts.
		ms -= (ms % msPerDay) * 2
	}
	return time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).Add(time.Duration(ms) * time.Millisecond)
}

// comArgs reicht Objekte als ihr IDispatch weiter (Move(ziel)).
func comArgs(args []any) []any {
	out := make([]any, len(args))
	for i, a := range args {
		if c, ok := a.(*comObj); ok {
			out[i] = c.d
		} else {
			out[i] = a
		}
	}
	return out
}

// comConn haelt Application und Namespace ueber die Requests hinweg. Pro
// Request entstehen frische Huellen mit eigener Arena.
type comConn struct {
	ol, ns *ole.IDispatch
	req    arena
}

func dialOutlook() (backend, error) {
	unk, err := oleutil.CreateObject("Outlook.Application")
	if err != nil {
		return nil, comErr(err)
	}
	defer unk.Release()
	ol, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, comErr(err)
	}
	v, err := oleutil.CallMethod(ol, "GetNamespace", "MAPI")
	if err != nil {
		ol.Release()
		return nil, comErr(err)
	}
	return &comConn{ol: ol, ns: v.ToIDispatch()}, nil
}

func (c *comConn) begin() (obj, obj) {
	return &comObj{c.ol, &c.req}, &comObj{c.ns, &c.req}
}

func (c *comConn) end() { c.req.release() }

func (c *comConn) close() {
	c.req.release()
	c.ns.Release()
	c.ol.Release()
}

// comError traegt den eigentlichen Fehlercode. Bei DISP_E_EXCEPTION steckt er
// in EXCEPINFO.scode, die Meldung in dessen Beschreibung — das ist der Text,
// den Outlook in seiner Sprache liefert ("The message was not found.").
type comError struct {
	code uint32
	msg  string
}

func (e *comError) Error() string { return e.msg }

func comErr(err error) error {
	var oe *ole.OleError
	if err == nil || !errors.As(err, &oe) {
		return err
	}
	code := uint32(oe.Code())
	if ei, ok := oe.SubError().(ole.EXCEPINFO); ok && ei.SCODE() != 0 {
		code = ei.SCODE()
	}
	msg := strings.TrimSpace(oe.Description())
	if msg == "" || msg == "<nil>" {
		msg = fmt.Sprintf("%s (0x%08X)", strings.TrimSpace(syscall.Errno(code).Error()), code)
	}
	return &comError{code, msg}
}

// isTransient: Outlook ist kurz nicht ansprechbar (beschaeftigt, frisch
// gestartet, neu gestartet). Darauf wird die Verbindung verworfen und der
// ganze Request wiederholt; alles andere ist ein fachlicher Fehler.
//
// 0x80010001 call rejected, 0x8001010A busy, 0x80010105 server fault,
// 0x80010108 disconnected (Outlook wurde neu gestartet), 0x80080005 exec
// failed, 0x800706BA/BE/BF RPC-Server weg bzw. Aufruf abgebrochen.
func isTransient(err error) bool {
	var ce *comError
	if !errors.As(err, &ce) {
		return false
	}
	switch ce.code {
	case 0x80010001, 0x8001010A, 0x80010105, 0x80010108, 0x80080005,
		0x800706BA, 0x800706BE, 0x800706BF:
		return true
	}
	return false
}
