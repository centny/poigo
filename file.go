package poigo

import (
	"github.com/Centny/jnigo"
)

type FileInputStream struct {
	Input *jnigo.Object
}

func (f *FileInputStream) Close() error {
	if f.Input == nil {
		return nil
	} else {
		return f.Input.CallVoid("close")
	}
}
func (f *FileInputStream) S() *jnigo.Object {
	return f.Input
}
func NewFileInputStream(path string) (*FileInputStream, error) {
	input, err := jnigo.GVM.NewAs("java.io.FileInputStream", "java.io.InputStream", path)
	if err == nil {
		return &FileInputStream{
			Input: input,
		}, nil
	} else {
		return nil, err
	}
}

type FileOutputStream struct {
	Output *jnigo.Object
}

func (f *FileOutputStream) Close() error {
	if f.Output == nil {
		return nil
	} else {
		return f.Output.CallVoid("close")
	}
}
func (f *FileOutputStream) S() *jnigo.Object {
	return f.Output
}
func NewFileOutputStream(path string) (*FileOutputStream, error) {
	output, err := jnigo.GVM.NewAs("java.io.FileOutputStream", "java.io.OutputStream", path)
	if err == nil {
		return &FileOutputStream{
			Output: output,
		}, nil
	} else {
		return nil, err
	}
}
