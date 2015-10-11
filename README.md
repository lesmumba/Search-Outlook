# Search-Outlook

This script uses the win32com module to search Microsoft Outlook foldersfor emails which match a set criteria.
Matching emails are saved as HTML pages to an appropriate folder for easy retrival labelled as:
  sender_name - message.Subject.html

A HTML page is also genearted which serves as an index file with links to emails in their respective locations.

## Note
You will need to install [pywin32](http://sourceforge.net/projects/pywin32/files/) for this to work.

Some links worth knowing about:
- [http://www.boddie.org.uk/python/COM.html](http://www.boddie.org.uk/python/COM.html)
- [https://msdn.microsoft.com/en-us/library/office/dn467914.aspx](https://msdn.microsoft.com/en-us/library/office/dn467914.aspx)
- [https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.aspx](https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.aspx)
- [https://msdn.microsoft.com/en-us/library/office/aa221870%28v=office.11%29.aspx](https://msdn.microsoft.com/en-us/library/office/aa221870%28v=office.11%29.aspx)
