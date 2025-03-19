const { DeleteObjectCommand, PutObjectCommand, GetObjectCommand, ListObjectsV2Command } = require("@aws-sdk/client-s3");
const { s3Client } = require("../s3Config");


// delete an object from the s3 bucket 
exports.deleteS3Object = async(bucketName,path) => {
    await s3Client.send(
        new DeleteObjectCommand({
            Bucket: bucketName,
            Key: path
        })
    );
};
// add or update an object into the s3 bucket
exports.addOrUpdateS3Object = async(bucketName,filePath,bufferFile) => {
    await s3Client.send(
        new PutObjectCommand({
            Bucket: bucketName,
            Key: filePath,
            Body: bufferFile,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // MIME type for Excel
        })
    )
}
// fetch a file from s3 bucket
exports.getObjectFromS3Bucket = async(bucketName,filePath) => {
    return await s3Client.send(
        new GetObjectCommand({
            Bucket: bucketName,
            Key: filePath,
        })
    );
}

exports.listObjectsFromS3Bucket = async(bucketName,filePath) => {
    return await s3Client.send(
        new ListObjectsV2Command({
            Bucket: bucketName,
            Prefix: filePath,
            Delimiter: '/', // Groups objects by folder
        })
    );
}