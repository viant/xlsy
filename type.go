package xlsy

import (
	"reflect"
)

func ensureStruct(t reflect.Type) reflect.Type {
	switch t.Kind() {
	case reflect.Struct:
		return t
	case reflect.Ptr:
		return ensureStruct(t.Elem())
	case reflect.Slice:
		return ensureStruct(t.Elem())
	}
	return nil
}

func ensureSlice(t reflect.Type) reflect.Type {
	switch t.Kind() {
	case reflect.Slice:
		return t
	case reflect.Ptr:
		return ensureSlice(t.Elem())
	case reflect.Struct:
		return nil
	}
	return nil
}
