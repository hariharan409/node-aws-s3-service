const s3Client = require("../common/aws-s3-config/s3Config");
const { ListObjectsV2Command } = require("@aws-sdk/client-s3");

exports.checkConfigOfS3 = async() => {
    try {
        const data = await s3Client.send(new ListObjectsV2Command({
            Bucket: "YOUR_BUCKET_NAME"
        }));
        console.log(data, "hii");
    } catch (error) {
        throw new Error(error.message || error);
    }
}