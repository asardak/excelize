package excelize

import "fmt"

func (f *File) AddDefinedName(sheet, begin, end, name string) {
	r := f.workbookReader()
	if r.DefinedNames == nil {
		r.DefinedNames = &xlsxDefinedNames{}
	}

	beginCol, beginRow := getCellColRow(begin)
	endCol, endRow := getCellColRow(end)
	r.DefinedNames.DefinedName = append(r.DefinedNames.DefinedName, xlsxDefinedName{
		Name: name,
		Data: fmt.Sprintf("'%s'!$%s$%s:$%s$%s", sheet, beginCol, beginRow, endCol, endRow),
	})
}

func (f *File) DeleteDefinedName(name string) {
	r := f.workbookReader()
	if r.DefinedNames == nil {
		r.DefinedNames = &xlsxDefinedNames{}
	}

	for i := 0; i < len(r.DefinedNames.DefinedName); i++ {
		if r.DefinedNames.DefinedName[i].Name == name {
			r.DefinedNames.DefinedName = append(r.DefinedNames.DefinedName[:i], r.DefinedNames.DefinedName[i+1:]...)
			return
		}
	}
}
