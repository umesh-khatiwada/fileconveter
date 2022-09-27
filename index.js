const express = require("express");
const dotenv = require("dotenv");
var sanitize = require("sanitize-filename");

const cors = require("cors");
//const cheerio = require("cheerio");
app = express();
const path = require("path");
dotenv.config({ path: "./config.env" });

const { exec } = require("child_process");
const PORT = process.env.PORT || 2000;

const { spawn } = require("child_process");

const $ = require("jquery");

const multer = require("multer");
const libre = require("libreoffice-convert");
var isJson = require("is-json");
const fs = require("fs");

app.use(express.json());
//app.use(bodyParser.urlencoded({ extended: false }));
//app.use(bodyParser.json());
app.use(cors());
app.set("view engine", "ejs");

app.use(express.static("public"));
var uploadsDir = __dirname + "/public/uploads";
var outputFilePath;

//Storage function
var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDir);
  },
  filename: function (req, file, cb) {
    cb(
      null,
      file.fieldname + "-" + Date.now() + path.extname(file.originalname)
    );
  },
});

//upload helper files

const docxtopdffilter2 = function (req, file, callback) {
  var ext = path.extname(file.originalname);
  console.log(ext);
  if (ext !== ".docx" && ext !== ".doc") {
    return callback("This Extension is not supported");
  }
  callback(null, true);
};

const docxtopdf2upload = multer({
  storage: storage,
  fileFilter: docxtopdffilter2,
}).single("file");
//= (pathoutput) =>

//Upload the pdf files
app.post("/uploaddocxtopdf", async (req, res) => {
  docxtopdf2upload(req, res, async function (err) {
    if (req.file) {
      outputFilePath = Date.now() + "output.pdf";
      const file = fs.readFileSync(req.file.path);
      libre.convert(file, ".pdf", undefined, (err, done) => {
        if (err) {
          console.log(err);
        }
        fs.writeFileSync(outputFilePath, done);
        var newPath = process.env.WEBSITENAMEDOCXTOPDF + outputFilePath;
        console.log(newPath);
        //return newPath;
        res.json({
          path: newPath,
        });
      });
    }
  });
});

//upload pdf file

const pdftodocxfilter = function (req, file, callback) {
  var ext = path.extname(file.originalname);
  if (ext !== ".pdf") {
    return callback("This Extension is not supported");
  }
  callback(null, true);
};

var pdftodocxstorage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "public/uploads");
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

var pdftodocxupload = multer({
  storage: pdftodocxstorage,
  fileFilter: pdftodocxfilter,
});

//pdf to Docx
app.post("/pdftodocx", pdftodocxupload.single("file"), (req, res) => {
  if (req.file) {
    var pdf = req.file.path;
    console.log(pdf);
    var basename = path.basename(req.file.path);
    //console.log(basename);

    basename = sanitize(basename);

    var htmlpath = `${basename.substr(0, basename.lastIndexOf("."))}.html`;
    // console.log(htmlpath);

    htmlpath = sanitize(htmlpath);

    var docxpath = `${basename.substr(0, basename.lastIndexOf("."))}.docx`;

    docxpath = sanitize(docxpath);

    //res.json({ ghello: "bmnb" });

    exec(
      `/Applications/LibreOffice.app/Contents/MacOS/soffice  --headless --convert-to html  ${pdf}`,

      //`soffice --headless --convert-to html --${uploadsDir} . ${pdf}`,
      (err, stdout, stderr) => {
        if (err) {
          // fs.unlinkSync(pdf);
          // fs.unlinkSync(htmlpath);
          // fs.unlinkSync(docxpath);
          res.send(err);
        }
        console.log(htmlpath.length);
        exec(
          `/Applications/LibreOffice.app/Contents/MacOS/soffice --convert-to docx:'MS Word 2007 XML' ${htmlpath}`,
          (err, stdout, stderr) => {
            if (err) {
              // fs.unlinkSync(pdf);
              // fs.unlinkSync(htmlpath);
              // fs.unlinkSync(docxpath);
              res.send(err);
            }
            console.log("output converted");
            //  console.log(docxpath.length);

            var newPath = process.env.WEBSITENAMEPDFTODOCX + docxpath;
            console.log(newPath);
            fs.unlinkSync(pdf);
            fs.unlinkSync(htmlpath);
            res.json({
              path: newPath,
            });

            // res.download(docxpath, (err) => {
            //   if (err) {
            //     // fs.unlinkSync(pdf);
            //     // fs.unlinkSync(htmlpath);
            //     // fs.unlinkSync(docxpath);
            //     res.send(err);
            //   }
            //   // fs.unlinkSync(pdf);
            //   // fs.unlinkSync(htmlpath);
            //   // fs.unlinkSync(docxpath);
            // });
          }
        );
      }
    );
  }
});

app.get("/downloadpdftodocx", (req, res) => {
  var pathoutput = req.query.path;
  console.log(pathoutput);
  var fullpath = path.join(__dirname, pathoutput);
  res.download(fullpath, (err) => {
    if (err) {
      fs.unlinkSync(pathoutput);
      res.send(err);
    }

    fs.unlinkSync(pathoutput);
  });
});

app.get("/download", (req, res) => {
  var pathoutput = req.query.path;
  console.log(pathoutput);
  var fullpath = path.join(__dirname, pathoutput);
  res.download(fullpath, (err) => {
    if (err) {
      fs.unlinkSync(fullpath);
      res.send(err);
    }
    fs.unlinkSync(fullpath);
  });
});

//app.use(require("./router/auth"));

app.listen(PORT, () => {
  console.log(`App is listening on Port ${PORT}`);
});
