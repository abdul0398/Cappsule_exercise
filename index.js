const express = require("express");
const app = express();
const path = require("path");
const reader = require("xlsx");
const fs = require("fs");
const fsPromises = require('fs').promises;
const multer = require("multer");
const upload = multer({ dest: "uploads/" });

app.use(express.urlencoded({ extended: true }));
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static(path.join(__dirname, "public")));

const file = reader.readFile("Data.xlsx");

const sheets = file.SheetNames;

const master = reader.utils.sheet_to_json(file.Sheets["Master"]);
const test = reader.utils.sheet_to_json(file.Sheets["Test"]);



let map = new Map();
for (let i = 0; i < master.length; i++) {
  const name = master[i]["Item name"]
    .toLowerCase()
    .split("")
    .filter((elem) => {
      return elem.toLowerCase() != elem.toUpperCase() || elem >= 0 || elem <= 9;
    })
    .join("");
  const id = master[i]["Product ID"];
  map.set(name, id);
}
for (let i = 0; i < test.length; i++) {
  const productName = test[i]["Item Name"]
    .toLowerCase()
    .split("")
    .filter((elem) => {
      return elem.toLowerCase() != elem.toUpperCase() || elem >= 0 || elem <= 9;
    })
    .join("");
  const productId = map.get(productName);

  test[i]["Product ID_Master"] = productId;
}



app.get("/", (req, res) => {
  res.render("index");
});



app.post("/fileUpload", upload.single("files"), async (req, res) => {
  const filename = req.file.originalname; // getting the original filename
  const oldPath = __dirname + `/uploads/` + req.file.filename;
  const newPath = __dirname + `/uploads/` + "file.xlsx";

  await fsPromises.rename(oldPath, newPath)

  console.log("before reading");
  const file = reader.readFile(`uploads/file.xlsx`); // reading the xlsx file
  console.log("After reading");
  // extracting different sheets from file
  const master = reader.utils.sheet_to_json(file.Sheets["Master"]);
  const test = reader.utils.sheet_to_json(file.Sheets["Test"]);

  // mapping the names with productId
  let map = new Map();
  for (let i = 0; i < master.length; i++) {
    const name = master[i]["Item name"]
      .toLowerCase()
      .split("")
      .filter((elem) => {
        return (
          elem.toLowerCase() != elem.toUpperCase() || elem >= 0 || elem <= 9
        );
      })
      .join("");
    const id = master[i]["Product ID"];
    map.set(name, id);
  }


  // updating the test to create Output sheet
  for (let i = 0; i < test.length; i++) {
    const productName = test[i]["Item Name"]
      .toLowerCase()
      .split("")
      .filter((elem) => {
        return (
          elem.toLowerCase() != elem.toUpperCase() || elem >= 0 || elem <= 9
        );
      })
      .join("");
    const productId = map.get(productName);

    test[i]["Product ID_Master"] = productId;
  }
  const ws = reader.utils.json_to_sheet(test);
  reader.utils.book_append_sheet(file,ws,"Outputt")
  // Writing to our file
  reader.writeFile(file,`uploads/file.xlsx`)
    res.render('download');


});



app.get('/download', function(req, res){
  const file = `${__dirname}/uploads/file.xlsx`;
  res.download(file, async function(err) {
    if(err){
      res.redirect('/');
      return;
    }
    let filePath = `${__dirname}/uploads/file.xlsx`; 
    await fsPromises.unlink(filePath);  
  });
});



app.listen(3000, () => {
  console.log("Server Started");
});
