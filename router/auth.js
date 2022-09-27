const express = require("express");
const router = express.Router();
const multer = require("multer");
const libre = require("libreoffice-convert");

const path = require("path");
const fs = require("fs");

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
  console.log("file.originalname");
  console.log(file.originalname);
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

//Upload the pdf files
app.post("/uploaddocxtopdf", (req, res) => {
  console.log(req.file);
  docxtopdf2upload(req, res, function (err) {
    if (req.file) {
      res.json({
        path: req.file.path,
      });
    }
  });
});

//Now converter the docx to pdf files

const docxtopdffilter = function (req, file, callback) {
  var ext = path.extname(file.originalname);

  console.log(file.originalname);

  if (ext !== ".docx" && ext !== ".doc") {
    return callback("This Extension is not supported");
  }
  callback(null, true);
};

const docxtopdfupload = multer({
  storage: storage,
  fileFilter: docxtopdffilter,
});

//Doc to pdf converter
router.post("/docxtopdf", (req, res) => {
  console.log(req);
  outputFilePath = Date.now() + "output.pdf";

  const file = fs.readFileSync(req.body.path);

  console.log(file);

  libre.convert(file, ".pdf", undefined, (err, done) => {
    if (err) {
      console.log(err);
    }
    fs.writeFileSync(outputFilePath, done);
    res.json({
      path: outputFilePath,
    });
  });
});

app.post("/docxtopdf", docxtopdfupload.single("file"), (req, res) => {
  console.log(req.file);

  if (req.file) {
    console.log(req.file.path);
    outputFilePath = Date.now() + "output.pdf";

    const file = fs.readFileSync(req.file.path);

    libre.convert(file, ".pdf", undefined, (err, done) => {
      if (err) {
        fs.unlinkSync(req.file.path);
        fs.unlinkSync(outputFilePath);
        res.send(`Error converting file: ${err}`);
      }
      console.log(done);

      // Here in done you have pdf file which you can save or transfer in another stream
      fs.writeFileSync(outputFilePath, done.toString());

      res.download(outputFilePath, (err) => {
        if (err) {
          fs.unlinkSync(req.file.path);
          fs.unlinkSync(outputFilePath);
          res.send("Unable to download the file");
        }
        fs.unlinkSync(req.file.path);
        fs.unlinkSync(outputFilePath);
      });
    });
  }
});

router.get("/download", (req, res) => {
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

module.exports = router;
