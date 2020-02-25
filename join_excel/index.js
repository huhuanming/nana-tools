const fs = require("fs");
const xlsx = require("node-xlsx").default;

const folderPath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/replace-name-18-25';

const files = fs.readdirSync(folderPath)

const joinSheets = [{
    name: "合并后表格",
    data: [],
}]
files.forEach((fileName) => {
    if (fileName.includes('.xlsx')) {
        const filPath = `${folderPath}\/${fileName}`
        const sheets = xlsx.parse(filPath);
        joinSheets[0].data = joinSheets[0].data.concat(sheets[0].data)
    }
})

fs.writeFileSync(`${folderPath}/join.xlsx`,  xlsx.build(joinSheets), { encoding: 'buffer' })
