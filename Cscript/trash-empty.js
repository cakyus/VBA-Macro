// =====================================================================
// @name         trash-empty
// @description  Empty Recycle Bin
// @usage        cscript.exe /nologo trash-empty.js
// =====================================================================

shell = new ActiveXObject('Shell.Application');
fso = new ActiveXObject('Scripting.FileSystemObject');
folder = shell.NameSpace(0xa);
folderItems = folder.Items();

for (i = 0; i < folderItems.Count; i++){
	folderItem = folderItems.Item(i);
	fso.GetStandardStream(2).WriteLine('> ' + folderItem.Name);
	if (folderItem.Type == 'File folder'){
		fso.DeleteFolder(folderItem.Path);
	} else {
		fso.DeleteFile(folderItem.Path);
	}
}
