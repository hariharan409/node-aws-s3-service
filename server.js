require("dotenv").config();
const express = require("express");
const { checkConfigOfS3 } = require("./aws/checkConfigOfS3");
const { generateChart } = require("./daily-engine-log/generateChart");
const app = express();

const port = 1997;

app.get("/",(req,res) => {
    res.send("i am ready to serve the data brother!");
});
// app.get("/upload-excel-to-s3",checkConfigOfS3);
generateChart();

app.listen(port,() => {
    console.log(`server is up and running on port ${port}`);
});