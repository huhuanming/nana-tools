const fs = require("fs")
const xlsx = require("node-xlsx").default


const filePath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/4aaee46d6b75ede5.xlsx'
const workSheets = xlsx.parse(filePath)

const names = []
workSheets[0].data.forEach((line, index) => {
    if(index > 0) {
        if(line[2] &&!names.includes(line[2])) {
            names.push(line[2])
        }
    }
})

fs.writeFileSync('names.txt', names.join('\n'), { encoding: 'utf8' })