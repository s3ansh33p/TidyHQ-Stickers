# Stickers

This is designed for ComSSA events, though anyone is welcome to use/mofidy this.
The word template is based on the `Avery_L7162GU` template, which is a 16 sticker per page template.

## How to use

1. Clone the repo
2. Modify the `ComSSATemplate.docx` file to your liking
3. Setup your credentials in the `.env` file, see `env setup` below
4. Edit the `thqhandler.js` file to your liking, notably with the group id
5. Run `node thqhandler.js` to fetch the contacts from TidyHQ
6. Run `node index.js` to generate the stickers

## Env setup
Create a `.env` file in the root of the project with the following contents:
```bash
CLIENT_ID=<your 64 character app id>
CLIENT_SECRET=<your 64 character app secret>
REDIRECT_URI=<your callback url from the app>
EMAIL=<your email from TidyHQ>
PASSWORD=<your account password from TidyHQ>
CLUB=<your domain prefix from TidyHQ>
```
For the CLIENT_ID, CLIENT_SECRET and REDIRECT_URI, you will need to create an app in TidyHQ, see [here](https://dev.tidyhq.com/oauth_applications). Read [here](https://dev.tidyhq.com/) for more info.