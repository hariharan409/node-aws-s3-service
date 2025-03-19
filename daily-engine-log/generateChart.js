const {spawn} = require("child_process");
const fs = require("fs");
const { getObjectFromS3Bucket } = require("../common/aws-s3-config/s3-operations/s3Operation");
const moment = require("moment-timezone");



exports.generateChart = async() => {
    const logDate = '01-03-2025';
    const bucketName = process.env.AWS_S3_EXCEL_UPLOAD_BUCKET_NAME;
    let subFolderPath = process.env.AWS_S3_EXCEL_UPLOAD_FOLDER_NAME + "" + moment(logDate,"DD-MM-YYYY").format("YYYY");
    let filePath = `${subFolderPath}/${moment(logDate,"DD-MM-YYYY").format("MMM").toLowerCase()}.xlsx`;
    // file has been available in the s3 bucket. fetching the workbook from s3 bucket
    const response = await getObjectFromS3Bucket(bucketName,filePath);

    // convert stream to buffer
    const fileBuffer = await streamToBuffer(response.Body);

    // spawn the python process
    const pythonProcess = spawn('python',["./python/add_chart_to_excel.py", logDate]);

    // Send the file buffer via stdin
    pythonProcess.stdin.write(fileBuffer);
    pythonProcess.stdin.end();

    // collect the processed buffer from Python
    let outputBuffer = Buffer.alloc(0);

    pythonProcess.stdout.on("data", (data) => {
        outputBuffer = Buffer.concat([outputBuffer, data]);
    });

    pythonProcess.stderr.on("data", (data) => {
        console.error(`stderr: ${data}`);
    });

    pythonProcess.on("close", (code) => {
        if (code === 0) {
            const outputFilePath = "C:/Users/Harihara.Dhamodaran/Documents/test.xlsx";
            // Save the modified buffer to an output Excel file
            fs.writeFileSync(outputFilePath, outputBuffer);
        } else {
            console.error(`Python script exited with code ${code}`);
        }
    });
}

/**
 * Converts a stream into a buffer
 * @param {Readable} stream - Readable stream
 * @returns {Promise<Buffer>}
 */
const streamToBuffer = async (stream) => {
    return new Promise((resolve, reject) => {
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        stream.on("end", () => resolve(Buffer.concat(chunks)));
        stream.on("error", reject);
    });
};