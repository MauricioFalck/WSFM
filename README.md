# WSFM
Copies the images from the Windows Spotlight service to be used as wallpapers

This small program gets the images from the Windows Spotlight service deployed locally and copies to a specific folder as jpg. 

Many improvements still pending, such as:
- configure the destination path (now it is fixed)
- log sent to Event Viewer, but the creation of the Application branch needs to be done manually
- transform the application in a service that will pull the files every xx hours
