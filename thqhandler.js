const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");
dotenv.config();

const { authorize, authorizeWithPassword, requestAccessToken } = require("../module/src/OAuth.js");
const TidyHQ = require("../module/index.js");

// get environment variables
const clientID = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const redirectURI = process.env.REDIRECT_URI;
const email_address = process.env.EMAIL;
const password = process.env.PASSWORD;
const club = process.env.CLUB;

let client;

async function auth() {
    let auth = await authorizeWithPassword(clientID, clientSecret, email_address, password, club);
    // console.log(auth)
    client = new TidyHQ(auth);
}

async function getGroup() {
    let group = await client.Groups.getGroupByName("2023-International Sundowner-28-08-MailingList");
    // let group = await client.Groups.getGroupByName("Current Members");
    let contacts = await client.Contacts.getContactsInGroup(group.id);
    let names = [];
    for (let i = 0; i < contacts.length; i++) {
        let contact = contacts[i];
        // console.log((contact.first_name).trim() + " " + (contact.last_name).trim());
        console.log(contact)
        names.push((contact.first_name).trim() + " " + (contact.last_name).trim());
    }
    console.log(names.length);
    // write to contacts.json
    fs.writeFileSync(path.resolve(__dirname, "target", "contacts.json"), JSON.stringify(names));
}

async function main() {
    await auth();
    await getGroup();
}

main();
