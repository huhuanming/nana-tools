const fs = require("fs")
const xlsx = require("node-xlsx").default


const filePath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/推月历史客户数据（新）-去潘朵拉.xlsx'
const workSheets = xlsx.parse(filePath)


const header = [workSheets[0].data[0]]
const sheetData = workSheets[0].data.slice(1, workSheets[0].data.length)
const sheetName = workSheets[0].name

const maxLength = parseInt(sheetData.length / 2000) + 1
console.log(`数据：${sheetData.length} 条`)

for (let index = 0; index < maxLength; index++) {
    const buffer = xlsx.build([{
        name: sheetName,
        data: header.concat(sheetData.slice(index * 2000, (index+1) * 2000)),
    }])
    fs.writeFileSync(`${index+1}-split-${index * 2000 + 1}-${(index+1) * 2000}.xlsx`, buffer, "binary");
}

