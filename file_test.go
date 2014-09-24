package poigo

import (
	"testing"
)

func tTestFile(t *testing.T) {
	NewFileInputStream("/tmsdf/dks")
	NewFileOutputStream("/tmsdf/sdffss")
	fis := FileInputStream{}
	fis.Close()
	fos := FileOutputStream{}
	fos.Close()
}
