# Productboard -> Google Sheets import
This is an example script for importing features with their custom fields from Productboard to Google sheets document. It is meant to showcase usage of Productboard [public API](https://developer.productboard.com).

## Recording
![Screen Recording 2022-05-24 at 14 51 24](https://user-images.githubusercontent.com/26506982/170039872-86041c07-11a5-4ba9-9d0b-c128ed1d7920.gif)
## How to use
1. Install [clasp](https://github.com/google/clasp)
2. `git clone https://github.com/productboard/pb-google-sheets-import` 
3. `clasp login`
4. `clasp create --type sheets` - note the created document link
5. `clasp push` 
7. `clasp deploy` 
8. Open the document created in step 4.
9. Wait for `Productboard` submenu to appear in the top menu.
10. Select `Productboard -> Import features`, authorize, add [Productboard public API token](https://developer.productboard.com/#section/Authentication) and hit `Ok`

For more information on development of Google AppScripts locally read [clasp documentation](https://developers.google.com/apps-script/guides/clasp).
