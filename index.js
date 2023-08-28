const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { exec } = require("child_process");

const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
    path.resolve(__dirname, "TemplateV4.docx"),
    "binary"
);

// read names from contacts.json 
let names = JSON.parse(fs.readFileSync(path.resolve(__dirname, "target", "contacts.json"), "utf8"));
// some bug I can't be bothered to fix so we're adding 16 blanks
for (let i = 0; i < 16; i++) {
    names.unshift(" ");
}

async function clearOld () {
    // clear all .ps1 and .docx files in target folder
    await fs.readdir(path.resolve(__dirname, "target"), (err, files) => {
        // if (err) console.log(err);
        for (const file of files) {
            if (file.endsWith(".ps1") || file.endsWith(".docx")) {
                // check not merge.ps1
                if (file == "merge.ps1") continue;
                if (file == "output.docx") continue;
                if (file.startsWith("blank")) continue;
                fs.unlink(path.join(path.resolve(__dirname, "target"), file), err => {
                    // if (err) console.log(err);
                });
            }
        }
    });
    // wait a second
    await new Promise(resolve => setTimeout(resolve, 2000));
}

async function main() {
    // build powershell script to remove images from docx if they are not needed (no data at index)
    await clearOld();
    let removeImages = "\n";
    let fname = "";

    // chunk array into 16
    for (let chunkIdx = 0; chunkIdx < names.length; chunkIdx += 16) {
        let chunk = names.slice(chunkIdx, chunkIdx + 16);
        let chunkId = chunkIdx / 16;
        fname = `part${chunkId}.docx`;
        console.log("Working on chunk: ", chunkId, " with length: ", chunk.length);
        
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        let changes = {};

        for (let i = 0; i < 16; i++) {
            // see if there is data at index
            let idx = i + chunkIdx;
            if (names[idx] == undefined) {
                // if not, remove image
                // powershell
                // if last chunk
                if (chunkId == Math.ceil(names.length / 16) - 1) {
                    removeImages += `if ($obj.AlternativeText -eq "${i+1}") {$obj.Delete();};\n`;
                }
                changes[`name${i+1}`] = " "; 
            } else {
                changes[`name${i+1}`] = names[idx];
            }
        }

        console.log("Changes: ", changes);

        doc.render(changes);

        const buf = doc.getZip().generate({
            type: "nodebuffer",
            // compression: DEFLATE adds a compression step.
            // For a 50MB output document, expect 500ms additional CPU time
            compression: "DEFLATE",
        });

        // buf is a nodejs Buffer, you can either write it to a
        // file or res.send it with express for example.
        await fs.writeFileSync(path.resolve(__dirname, "target", fname), buf);

    };
    let script = `$MSWord = New-Object -com word.application;
$MSWord.Visible = $false;
$WorkFile = "${path.resolve(__dirname, "target", fname)}";
$doc = $MSWord.Documents.Open($WorkFile);
foreach($obj in $doc.InlineShapes){${removeImages}}
$MSWord.ActiveDocument.Save();
$MSWord.ActiveDocument.Close();
$MSWord.Quit();
Write-OutPut "Done with ${fname}";`;

    // write script to file
    await fs.writeFileSync(path.resolve(__dirname, "target", `cleanup.ps1`), script);
    await execScript();
};
async function execWithPromise(command) {
    return new Promise((resolve, reject) => {
        exec(command, (error, stdout, stderr) => {
            if (error) {
                reject(error);
                return;
            }
            if (stderr) {
                reject(stderr);
                return;
            }
            console.log(stdout);
            resolve(stdout);
        });
    });
}
// powershell execution
async function execScript() {
    let outputFileLength = Math.ceil(names.length / 16);
    // run powershell script
    console.log("Skipping powershell script: cleanup.ps1");
    // console.log(`Running powershell script: ${path.resolve(__dirname, "target", `cleanup.ps1`)}`);
    // await execWithPromise(`powershell.exe -ExecutionPolicy Bypass -File ${path.resolve(__dirname, "target", `cleanup.ps1`)}`)
    // finally execute merge command

    // get length of output files
    console.log("Creating merge script for ", outputFileLength, " files");
    let mergeScript = `
$path = "${path.resolve(__dirname, "target")}";
Write-Host "Path: $path"
$docArray = (0..${outputFileLength-1} | ForEach-Object { Join-Path $path "part$_.docx" });

Write-Host "Creating Word COM object"
$word = New-Object -ComObject Word.Application
$word.Visible = $false

Write-Host ("Merging {0} documents" -f $docArray.Count)
$doc   = $word.Documents.Open($docArray[0])
$range = $doc.Range()
Write-Host ("Opened {0}" -f $docArray[0]);

# loop through each
foreach ($docPath in $docArray[1..${outputFileLength-1}]) {
    Write-Host ("Inserting {0}" -f $docPath)
    # move to the end of the first document
    $range.Collapse(0) # wdCollapseEnd see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcollapsedirection
    $range.InsertBreak(6)  # wdLineBreak   see https://learn.microsoft.com/en-us/office/vba/api/word.wdbreaktype
    # Note wdPageBreak had some weird behavior
    $range.InsertFile($docPath)
}

Write-Host "Saving document"
$doc.SaveAs($path + "\\output.docx")

# quit Word and cleanup the used COM objects
$word.Quit()
Write-Host "Closing Word and clearing COM objects"

foreach ($d in $range) {
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($d)
}
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-OutPut "Done with merge.ps1";`
    await fs.writeFileSync(path.resolve(__dirname, "target", `merge.ps1`), mergeScript);
    console.log(`Running powershell script: ${path.resolve(__dirname, "target", `merge.ps1`)}`);
    await execWithPromise(`powershell.exe -ExecutionPolicy Bypass -File ${path.resolve(__dirname, "target", `merge.ps1`)}`)
    console.log("Clearing unneeded files");
    // clear unneeded files
    await clearOld();
    console.log("Done");  
};

main();
// clearOld();

