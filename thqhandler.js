const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");
dotenv.config();

const { authorizeWithPassword } = require("./node_modules/tidyhq/src/OAuth");
const TidyHQ = require("tidyhq");

// get environment variables
const clientID = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const email_address = process.env.AUTH_USER;
const password = process.env.AUTH_PASS;
const club = process.env.DOMAIN_PREFIX;
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;

let client;

async function auth() {
    let auth = await authorizeWithPassword(clientID, clientSecret, email_address, password, club);
    console.log(auth);
    client = new TidyHQ(auth);
}

async function getGroup() {
    client = new TidyHQ(ACCESS_TOKEN);
    // let group = await client.Groups.getGroupByName("2023-International Sundowner");
    // let group = await client.Groups.getGroupByName("Current Members");
    // let contacts = await client.Contacts.getContactsInGroup(group.id);
    let contacts = await client.Contacts.getContactsInGroup(180107);
    contacts = contacts.data;
    let names = [];
    for (let i = 0; i < contacts.length; i++) {
        let contact = contacts[i];
        // console.log((contact.first_name).trim() + " " + (contact.last_name).trim());
        names.push((contact.first_name).trim() + " " + (contact.last_name).trim());
    }
    console.log(names.length);
    // write to contacts.json
    fs.writeFileSync(path.resolve(__dirname, "target", "contacts.json"), JSON.stringify(names));
}

async function main() {
    // await auth();
    await getGroup();
}

main();
