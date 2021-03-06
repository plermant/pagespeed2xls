# pagespeed2xls
Export the result of Google pagespeed insights API into an Excel spreadsheet. 
This can be useful when you need pagespeed results for many URLs (= batch capability) or want to access the results for presentation purposes, or when offline

It leverages the xlwt python module you'll need to install first. 
The main program was tested on python 2.7 and will likely break on python 3

Before you can use it, you'll need an API key of type 'browser' to access the pagespeed API. Follow instructions listed at https://support.google.com/cloud/answer/6158862 to acquire a key. You should also enable the PageSpeed Insight API from the Google Developers Console 

To validate your key, paste the following into your browser, and it should return a valid json payload (responseCode=200):
https://www.googleapis.com/pagespeedonline/v2/runPagespeed?url=http://code.google.com/speed/page-speed/&key=yourAPIKey

Once you have a valid API key, you can use the program in 2 ways, as listed in 'pagespeed2xls.py -h':

1) Interactive, if you want to only test one URL at a time.
Type 'pagespeed2xls.py', then 
type in your API key upon request, then
type in your URL 

 2) Input file.
This file has your API key as the first line, and then one URL per line. 
e.g. type 'pagespeed2xls.py -i \<inputfile\> -o \<outputfile.xls\>'

If you chose to rename the default output file, make sure you add extension '.xls' so it can be opened automatically by Excel.

The output is an excel spreadsheet that captures all (at least most important) the info reported by the pagespeed api. Only difference is that under the 'browser caching' column, I only report assets that are not cached, whereas the UI reports short caching times. Three tabs are created: Mobile Speed, Mobile Usability, Desktop Speed.

For trouble shooting purposes, the json payloads for mobile and desktop are also saved to temporary files
