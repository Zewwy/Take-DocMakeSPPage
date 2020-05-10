# Take-DocMakeSPPage
This script takes a Word Doc and converts it to HTML, uploads the images to a SharePoint Doc Library, then links the images in the generate HTML.

Most user input is well validated, Word doc(x) and folder paths are well validated, pre-reqs are checked decently well, only things I didn't code for were existing files and such, but even the destination folder on SharePoint is validated.

This doesn't actually create a SharePoint page it generates the proper HTML that you can paste in the "edit source code" of a SharePoint page which is HTML, with the proper img reference tags the pasted code will load just fine with the images sourced relatively on the destination SharePoint Page as long as the page being edited is on the site refernced in the script were the images are uploaded to.
