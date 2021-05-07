# Copy-MediaFilesFromMobile
Copies files from MTP connected devices to local drive

Sharing this mainly as an example of how to access MTP connected devices via Powershell. I remember that I have spent quite some time looking for such an example on the web.

### What is it?

I wanted to download images/videos from few mobile phones periodically without removing them from the device. The problem was that MTP propocol is pretty slow and when I was accessing a "camera" directory via Explorer it was taking ages to load the list of the files, find the last photo which was downloaded previosly and copy over the new ones. This script automates this process.

### What does it do exactly?

* Loads config file
* Scans for devices listed in config
* When device is found scans the list of content in the directory specified in the config
* Checks how many files are going to be downloaded
* Propmpts for confirmation before download
* Downloading files to the directory specified in the config
* Once it is finished asks whether to update the config with the latest file timestamp

### Troubleshooting

Initial steps of the script can be very slow if you have a lot of files in the directory on your device.

Add `-Verbose` switch to see the debug otput
