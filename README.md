# GostCode

This program inserts code in text form to Microsoft Word documents. This allows generating Gost documents (e.g. "Program source"). Original template might be in .doc and .docx, hence the program uses Microsoft.Office.Interop.Word to ensure that original template is not altered.

## Usage

 - Prepare template.doc. Insert "\<Inject\>" and "\<ListCount\>" text in the document. Later it will be replaced with the actual content.
 - Prepare file list yml. See the example in "example" directory.
 - Simply call the tool using the following command line:
```GostCode.exe template.doc out.doc C:\ProjectDirectory FileList.yml```

## Building

This project has only been tested on Windows. We don't even suppose that Linux version works.

```dotnet build GostCode.csproj```

# License

MIT license on the project itself.
Microsoft Word is a commercial product of Microsoft Corporation.
