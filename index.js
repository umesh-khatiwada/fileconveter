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
const { readdirSync, rename } = require("fs");

app.use(express.json());
//app.use(bodyParser.urlencoded({ extended: false }));
//app.use(bodyParser.json());
app.use(cors());
app.set("view engine", "ejs");

app.use(express.static("public"));
var uploadsDir = __dirname + "/public/uploads";
var outputFilePath;
var maxSize = 10000 * 1024 * 1024;

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
//IMAGE FILTER
const imageFilter = function (req, file, callback) {
  var ext = path.extname(file.originalname);
  if (
    ext !== ".png" &&
    ext !== ".jpg" &&
    ext !== ".jpeg" &&
    ext !== ".bmp" &&
    ext !== ".tiff" &&
    ext !== ".gif" &&
    ext !== ".wmf" &&
    ext !== ".pdf"
  ) {
    return callback("This Extension is not supported");
  }
  callback(null, true);
};

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
        var newPath = process.env.WEBSITENAME + outputFilePath;
        console.log(newPath);
        //return newPath;
        res.json({
          path: newPath,
          ext: ".pdf",
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

    var newFilepath = pdf;
    newFilepath = newFilepath.replace(/\s/g, "").toLowerCase();

    //rename
    const oldPath = pdf;

    // lowercasing the filename
    const newPathh = newFilepath;

    // Rename file
    rename(oldPath, newPathh, (err) => console.log(err));
    var basename = path.basename(newFilepath);
    //console.log(basename);

    basename = sanitize(basename);

    var htmlpath = `${basename.substr(0, basename.lastIndexOf("."))}.html`;
    // console.log(htmlpath);

    htmlpath = sanitize(htmlpath);

    var docxpath = `${basename.substr(0, basename.lastIndexOf("."))}.docx`;

    docxpath = sanitize(docxpath);

    exec(
      `$SOFFICE  --headless --convert-to html  ${newFilepath}`,

      //`soffice --headless --convert-to html --${uploadsDir} . ${pdf}`,
      (err, stdout, stderr) => {
        if (err) {
          fs.unlinkSync(newFilepath);
          // fs.unlinkSync(htmlpath);
          // fs.unlinkSync(docxpath);
          return res.send(err);
        }
        console.log(htmlpath.length);
        exec(
          `$SOFFICE --convert-to docx:'MS Word 2007 XML' ${htmlpath}`,
          (err, stdout, stderr) => {
            if (err) {
              fs.unlinkSync(pdf);
              fs.unlinkSync(htmlpath);
              fs.unlinkSync(docxpath);
              return res.send(err);
            }
            console.log("output converted");

            var newPath = process.env.WEBSITENAME + docxpath;
            console.log(newPath);
            fs.unlinkSync(newFilepath);
            fs.unlinkSync(htmlpath);
            res.json({
              path: newPath,
            });
          }
        );
      }
    );
  }
});

//txt to doc
const txttodocfilter = function (req, file, callback) {
  var ext = path.extname(file.originalname);
  if (ext !== ".txt") {
    return callback("This Extension is not supported");
  }
  callback(null, true);
};

var txttodocupload = multer({
  storage: storage,
  limits: { fileSize: maxSize },
  fileFilter: txttodocfilter,
});

app.post("/txttodoc", txttodocupload.single("file"), (req, res) => {
  if (req.file) {
    var file = req.file.path;

    var output = Date.now() + "output.doc";

    var txt = fs.readFileSync(file);

    fs.writeFileSync(output, txt);
    var newPath = process.env.WEBSITENAME + output;
    console.log(newPath);
    res.json({
      path: newPath,
    });
  }
});

//image to pdf

var imagetopdfupload = multer({
  storage: storage,
  fileFilter: imageFilter,
}).array("file", 100);

app.post("/uploadimagetopdf", (req, res) => {
  imagetopdfupload(req, res, function (err) {
    if (err) {
      return res.end("Error uploading file.");
    }
    var list = "";
    req.files.forEach((file) => {
      list += `${file.path}`;
      list += " ";
    });

    outputFilePath = Date.now() + "output.pdf";
    console.log(outputFilePath);

    exec(`img2pdf ${list} -o ${outputFilePath}`, (err, stdout, stderr) => {
      if (err) {
        return res.json({
          error: "some error takes place",
        });
      }

      var newPath = process.env.WEBSITENAME + outputFilePath;

      res.json({
        path: newPath,
      });
    });
  });
});

//download
app.get("/download", (req, res) => {
  var pathoutput = req.query.path;
  console.log(pathoutput);
  var fullpath = path.join(__dirname, pathoutput);
  res.download(fullpath, (err) => {
    if (err) {
      fs.unlinkSync(fullpath);
      return res.send(err);
    }
    fs.unlinkSync(fullpath);
  });
});

app.listen(PORT, () => {
  console.log(`App is listening on Port ${PORT}`);
});
