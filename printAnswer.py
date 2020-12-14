def printFile(filename):
    import tempfile
    import win32api
    import win32print
    open(filename, "r")
    win32api.ShellExecute(
        0,
        "print",
        filename,
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )

printFile('answer.docx')