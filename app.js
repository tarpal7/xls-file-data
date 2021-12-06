const express = require("express");
const app = express();
const bodyParser = require("body-parser");
const multer = require("multer");
const xlstojson = require("xls-to-json-lc");
const xlsxtojson = require("xlsx-to-json-lc");
const csvtojson = require("csvtojson");
const timeout = require("connect-timeout"); 
const fs = require("fs");
const dotenv = require("dotenv");
dotenv.config({ path: "./.env" });

const port = process.env.PORT || 3000;
app.use(timeout(180000));
app.use(bodyParser.json());
app.use(
  bodyParser.urlencoded({
    limit: "25mb",
    parameterLimit: 100000,
    extended: false,
  })
);

const storage = multer.diskStorage({
  //multers disk storage settings
  destination: function (req, file, cb) {
    cb(null, "./uploads/");
  },
  filename: function (req, file, cb) {
    var datetimestamp = Date.now();
    cb(null, file.originalname);
  },
});

const upload = multer({
  //multer settings
  storage: storage,
  fileFilter: function (req, file, callback) {
    //file filter
    if (
      ["xls", "xlsx", "csv"].indexOf(
        file.originalname.split(".")[file.originalname.split(".").length - 1]
      ) === -1
    ) {
      return callback(new Error("Wrong extension type"));
    }
    callback(null, true);
  },
}).single("file");

/** API path that will upload the files */
app.post("/upload", function (req, res) {
  let exceltojson;
  upload(req, res, function (err) {
    if (err) {
      res.json({ error_code: 1, err_desc: err });
      return;
    }
    /** Multer gives us file info in req.file object */
    if (!req.file) {
      res.json({ error_code: 1, err_desc: "No file passed" });
      return;
    }
    /** Check the extension of the incoming file and
     *  use the appropriate module
     */
    const extFile =
      req.file.originalname.split(".")[
        req.file.originalname.split(".").length - 1
      ];
    if (extFile === "xlsx") {
      exceltojson = xlsxtojson;
    } else if (extFile === "xls") {
      exceltojson = xlstojson;
    } else if (extFile === "csv") {
      csvtojson()
        // {
        // delimiter: [";"],
        // }
        .fromFile("./uploads/" + req.file.originalname)
        .then((result) => {
          res.json({
            error_code: 0,
            err_desc: null,
            data: result,
          });
        })
        .catch((err) => {
          // log error if any
          console.log(err);
        });
    }
    if (extFile === "xls" || extFile === "xlsx") {
      try {
        exceltojson(
          {
            input: req.file.path,
            output: "./uploads/data.json",
            lowerCaseHeaders: true,
          },
          function (err, result) {
            if (err) {
              return res.json({ error_code: 1, err_desc: err, data: null });
            }
            res.json({ error_code: 0, err_desc: null, data: result });
          }
        );
      } catch (e) {
        res.json({ error_code: 1, err_desc: "Corupted excel file" });
      }
    }
  });
});

app.get("/", function (req, res) {
  res.sendFile(__dirname + "/index.html");
});

app.get("/file", function (req, res) {
  const parametrs = req.query;
  const fileName = parametrs.name;

  const outObject = {};
  outObject.file = fileName;

  const extFile = fileName.split(".")[fileName.split(".").length - 1];
  if (extFile === "xlsx") {
    exceltojson = xlsxtojson;
  } else if (extFile === "xls") {
    exceltojson = xlstojson;
  } else if (extFile === "csv") {
    csvtojson()
      // {
      // delimiter: [";"],
      // }
      .fromFile("./uploads/" + fileName)
      .then((result) => {
        res.json({
          error_code: 0,
          err_desc: null,
          data: result,
        });
      })
      .catch((err) => {
        // log error if any
        console.log(err);
      });
  }

  if (extFile === "xls" || extFile === "xlsx") {
    try {
      exceltojson(
        {
          input: "./uploads/" + fileName,
          output: null, //since we don't need output.json
          lowerCaseHeaders: true,
        },
        function (err, result) {
          if (err) {
            return res.json({
              error_code: 1,
              err_desc: err,
              data: null,
            });
          }
          res.json({ error_code: 0, err_desc: null, data: result });
        }
      );
    } catch (e) {
      res.json({ error_code: 1, err_desc: "Corupted excel file" });
    }
  }
});

app.get("/check-data", function (req, res) {
  const result = {};
  result["success"] = true;
  result["message"] = "";
  result["data"] = {};

  const parametrs = req.query;

  console.log(parametrs)
  const dni = parametrs.dni;
  const phone = parametrs.phone;
console.log(dni, phone)
  try {
    const data = fs.readFileSync("./uploads/data.json", "utf8");
    const objs = JSON.parse(data);

    const findData = objs.filter((item) => {
      return item.dni === dni && item.phone === phone;
    })[0];
    if (findData) {
      result.data = findData;
    } else {
      result["success"] = false;
      result["message"] = "object not found";
      result["data"] = {};
    }
  } catch (err) {
    result.success = false;
    result.message = String(err);
  }
  res.json(result);
});


app.get("/pdf", function (req, res) {
  const parametrs = req.query;
  const filePath = parametrs.path;

  fs.readFile(__dirname + filePath, function (err, data) {
    res.contentType("application/pdf");
    res.send(data);
  });
});

app.get("/beneficios", function (req, res) {
  var fs = require("fs");
  var path = require("path");

  var getFiles = function (dir, files_) {
    files_ = files_ || [];
    filesShortPath = [];
    var files = fs.readdirSync(dir);
    for (var i in files) {
      var name =  path.join(dir,files[i]);
      if (fs.statSync(name).isDirectory()) {
        getFiles(name, files_);
      } else {
        files_.push(name);
        filesShortPath.push;
      }
    }
    return files_;
  };

  const dirFiles = path.join(__dirname, 'uploads', 'beneficios');
  
  const allFiles = getFiles(dirFiles);
  console.log(allFiles);
    res.send(allFiles);

});

app.listen(port, function () {
  console.log(`------running on ${port}...`);
});
