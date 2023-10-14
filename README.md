# About "ics2csv4Win"
This is a batch-exporter for online ics calendars. It will download and transfer the selected appointments as character separated values (csv) without the need of installing extra packages, frameworks, libraries etc. - Just windows batch processing!

# Background | Intention
It was intentionally created for users with windows machines - e.g. for employees from associations or organizations to further post processing their online calendars e.g. in excel, word etc.

# Configuration
1. Download source-code zip from latest release
2. Unpack
3. Edit the file "ics2.csv.bat"
4. Change the following param value to your online ics resource:
```
set "cal_url=https://<domain>/<path_to_ics_file>"
```
5. Save file and close editor

# Usage
1. Run the exporter by double click onto file "ics2csv.bat"
2. Type in the desired start date and confirm with the return key
3. Type in the desired end date and confirm with the return key
4. Download starts...
5. After the download has been finsished, the exported appointments will be automatically opened in a separate editor window
6. The exported appointments are separated by TABs and can be easily transferred into excel by copy & paste
7. While the downloading and exporting process is active, there will be created some temporary files which are automatically removed again after closing the extra editor window.

# Miscellaneous 
* It would be very nice if someone is motivated to port this code to bash, node.js or whatever else - please get in touch!
* This is tested under Windows 11 - But it should be also working with Windows 10 since curl[^1] (needed for downloading) is delevered by default.curl

[^1]: https://curl.se/