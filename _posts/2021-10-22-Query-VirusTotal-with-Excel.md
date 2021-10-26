---
published: false
---
## Problem

You have monotonous tasks at work that require a lot clicking, copying, and pasting content with a browser. Surely there is a better way to do this. But if you're at a large, security-mided organization, then most likely you will not have permissions to run your \[insert favorite scripting language here\] in the environment. So what now?

The answer: Excel

## Solution!

Did you know Excel uses [Power Query M ](https://docs.microsoft.com/en-us/powerquery-m/) language to batch and query data? You can use Power Query to call an API for automation.

1. VirusTotal provides their v3 API. Just create a VT account and click on "API Key" under your profile picture. Copy this for future use.

### Easy way to create function


2. Open Excel > Data tab > 'Get & Transform Data' section > Get Data > From Other Sources > From Web > Advanced

3. Now enter VT URL of an example IOC you want to query using the API format (e.g., https://www.virustotal.com/api/v3/ip_addresses/xxx.xxx.xxx.xxx. You can get the full documentation for the API at https://developers.virustotal.com/reference.

4. Enter HTTP request header (x-apikey) with your API Key value in the input box. Then hit OK.

5. This will open a new worksheet with that IP address result. Open the result in the Power Query editor. This allows you to go through the JSON object tree. Once you find the specific field you want to reference, right click and choose 'Drill Down'. Go to Data tab > Show Data & Connections. This opens a side panel with the query you just created listed. Right click query > Edit..

6. This opens the Power Query editor window. Expand left-hand Queries side panel > right click query > Create Function.

7. This opens a new Power Query editor window. Now choose Advanced Editor. We will now create variables for the API Key and the IP address, like the following:

8. Save this as a function connection.

### Do it again

You can create a function with any field in the VirusTotal query JSON result. Then import these .odc PQ connection files to any workbook you want and then use the From Table/Range option in Data tab > Add Column tab > then Invoke Custom Function. This will apply your function to the values in the table. If you have a premium VT API Key, this will work right away and give you all the results you need. However, if you have a basic, freemium key, VirusTotal limits your API requests to 4/min -- meaning you will need to invoke the function after 15 seconds. I will address this in another post once I get it to work properly on my end.