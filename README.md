# gPdfMerge
GUI-based PDF Merger 

**Latest version:** v1.2.14, 07 July 2024

![image](https://github.com/fafalone/gPdfMerge/assets/7834493/0ff980bd-99f2-4f36-b698-1320119c4994)

gPdfMerge is a simple utility written in twinBASIC mainly for me to try out using pdfium.dll, Google's open source PDF utility. It will merge the selected PDFs together, either into a new file, or into the first file on the list. You can optionally specify which pages in each document should be merged, and if you're appending the first in list, where to insert them at. That's all there is to it right now; just a brief experiment before I make a 64bit version of a more complex PDF control.


**Updates:**

(Version 1.2 - 07 Jul 2024)\
-Adds a 'Search for pages' function which opens a new dialog you can use to search for a range of pages to merge:

![image](https://github.com/fafalone/gPdfMerge/assets/7834493/6370de38-ea8c-4f40-91cc-c895c1456e7d)

(Version 1.1)\
-Range entry textbox now only enabled if an item is selected.\
-Specifying a single pdf with page range to trim it now supported.

**Build notes:**\
The project is configured as follows:\
-The project's root folder should contain the .twinproj file, and folders `win32` and `win64` with the respective bitness version of pdfium.dll\
-The compiled project must have pdfium.dll in the same folder with the .exe (the build output is set to \win32 and \win64).


One thing of interest, all the VB6 samples didn't have save functionality. I wasn't positive on how to implement it at first; it uses a very VB/tB unfriendly method:

```c
typedef struct FPDF_FILEWRITE_ {

  int version;

  int (*WriteBlock)(struct FPDF_FILEWRITE_* pThis,
                    const void* pData,
                    unsigned long size);
} FPDF_FILEWRITE;

FPDF_EXPORT FPDF_BOOL FPDF_CALLCONV FPDF_SaveAsCopy(FPDF_DOCUMENT document,
                                                    FPDF_FILEWRITE* pFileWrite,
                                                    FPDF_DWORD flags);
```

Functions in UDTs defined like that isn't something you usually see. pdfium makes you write your own write function. Other languages weren't helpful here... it seems all the ones I looked at had a built in class that somehow worked with this layout. We're not so lucky. It looked like a simple function pointer to a routine with a file write would work, and it did with one caveat: Despite the DLL's 32bit export calling convention being _stdcall, this callback had to be _cdecl, which might explain why nobody had done it in VB6 before. While there were solutions for APIs, afaik it wasn't until 2021 CDecl functions became practical with The trick's VBCDeclFix, an amazing piece of work that actually seems to finish VB6's incomplete CDecl support. 

twinBASIC, of course, supports CDecl natively, so no hacks are needed.

The code opens for write the output using `CreateFile` prior to calling `FPDF_SaveAsCopy`, then fills in a copy of the UDT with `AddressOf WriteBlock` 

```vba
    Private Type FPDF_FILEWRITE
        version As Long
        WriteBlock As LongPtr
    End Type

            Dim tWrite As FPDF_FILEWRITE
            tWrite.version = 1
            tWrite.WriteBlock = AddressOf WriteBlock

    Private Function WriteBlock CDecl(ByVal pThis As LongPtr, ByVal pData As LongPtr, ByVal size As Long) As Long
        If hFileOut Then
            Dim cbRet As Long
            Return WriteFile(hFileOut, ByVal pData, size, cbRet, vbNullPtr)
        End If
    End Function
```

Et voilà.

------

### Thanks and Notes
 Developed using pdfium builds by pvginkel,\
 32bit: https://github.com/pvginkel/PdfiumBuild/blob/master/Builds/2018-04-08/Pdfium-x86-no_v8-no_xfa/pdfium.dll \
 64bit: https://github.com/pvginkel/PdfiumBuild/blob/master/Builds/2018-04-08/Pdfium-x64-no_v8-no_xfa/pdfium.dll \
 This is an older version; the exact version of pdfium used couldn't be determined (none released on date of builds), the closest match seems to be pdfium 3392. This was done because the 32bit DLL here was [known to work in VB6](https://www.vbforums.com/showthread.php?882115-PDF-Reader-ActiveX-Control) without issue; in future versions I'll test more recent builds of pdfium.\
 Headers used for declares: https://pdfium.googlesource.com/pdfium/+/refs/heads/chromium/3392/public/ \
 Merge routine based on pdfium-cli: https://github.com/klippa-app/pdfium-cli

Also thanks to tB Discord user mike webb for getting me interested in this and helping test the program while under development.
 

 **Command line**\
 Command line usage:
 
Merge: gPdfMerge.exe /i "C:\...\Input1.pdf" "C:\...\Input2.pdf" /o "C:\path\Output.pdf"

Append: gPdfMerge.exe /i "C:\...\Input1.pdf" "C:\...\Input2.pdf"

Append or merge with ranges and/or insert idx not support via command line in v1.0


 
