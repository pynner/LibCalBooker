LibCal Booker
=========
LibCal booker is an open source tool to make booking LibCal study rooms easier. 

Currently configured for Lakehead University study rooms, but can be forked and adjusted for other schools using LibCal software.


Prerequisites  
----------
* Google Chrome or Firefox installed
* Python 2.7 
* python set as an environment variable

Enable Auto-Confirm Feature
---
1. Use [this wizard](https://console.developers.google.com/start/api?id=gmail) to create or select a project in the Google Developers Console and automatically turn on the API. Click **Continue**, then **Go to credentials**.
1. On the **Add credentials to your project** page, click the **Cancel** button.
1. At the top of the page, select the **OAuth consent screen** tab. Select an **Email address**, enter a **Product name** if not already set, and click the **Save** button.
1. Select the **Credentials** tab, click the **Create credentials** button and select **OAuth client ID**.
1. Select the application type **Other**, enter the name "LibCalBooker Email Confirm", and click the **Create** button.
1. Click **OK** to dismiss the resulting dialog.
1. Click the ![file_download](https://i.imgur.com/qaBisiO.png "Download") (Download JSON) button to the right of the client ID.
1. Move this file to your working directory and rename it `client_secret.json`.

Run 
-----
* run `python libCal.py` from the project root
