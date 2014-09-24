JNI-GO
======
> library for POI by golang through JNI

===
#Install
* install jdk (install mingw with gcc on window)
* go get:

linux/unix/osx:

 ```
export CGO_CFLAGS="-I$JAVA_HOME/include/ -I$JAVA_HOME/include/<darwin/linux>"

export CGO_LDFLAGS="-L$JAVA_HOME/jre/lib/server -ljvm"

go get github.com/Centny/jnigo
go get github.com/Centny/poigo
```

win32:

```
set CGO_CFLAGS=-I%JAVA_HOME%\include -I%JAVA_HOME%\include\win32

set CGO_LDFLAGS=-L%JAVA_HOME%\lib -ljvm

go get github.com/Centny/jnigo
go get github.com/Centny/poigo	
```


#Example

```go

	Init("<class path for poi>")
	input, err := NewFileInputStream("t.xlsx")
	if err != nil {
		return
	}
	defer input.Close()
	wb, err := OpenXSSFWorkbook(input)
	if err != nil {
		return
	}
	sheet, err := wb.SheetAt(0)
	if err != nil {
		return
	}
	err = sheet.Loop(func(r *Row) {
		fmt.Println("xxxxxx000000xxxxx")
		fmt.Println(r)
		fmt.Println(r.PhysicalNumberOfCells())
		r.Loop(func(c *Cell) {
			fmt.Println(c.String())
		})
	})
	if err != nil {
		return
	}
```