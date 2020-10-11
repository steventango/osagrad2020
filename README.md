# osagrad2020
Google App Scripts used to generate the banquet slideshow and automate compliment cards Google Slides for OSA's Class of 2020.

## generate_banquet_slideshow.gs
  - Used to generate the banquet Google Slides slideshow for the Class of 2020 from our Google Form spreadsheet data using a Google Slides template.
## automate_compliment_cards.gs
  - Used to automatically convert compliment cards from our Google Form spreadsheet data to printable fancy Compliment Cards using a Google Slides template.

## Usage Notes
  - Replace [REDACTED] with relevant Google Sheets/Slides IDs.
  - Make sure your Google Sheets columns match with the code.
  - You need to enable the [Google Slides API](https://developers.google.com/slides/quickstart/apps-script)
  - GET_LAYOUTS() is a helper function, and can be used to find the IDs of Placeholders.
  - For generate_banquet_slideshow(), make sure to run the UNSHARE() function to maintain file integrity.