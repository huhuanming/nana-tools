const fs = require("fs");
const xlsx = require("node-xlsx").default;

// const sourceFilePath = process.argv.pop();
// const folderPath = process.argv.pop();

const sourceFilePath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/过滤替换公司名/推月聚流-关联医院.xlsx';
const folderPath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/过滤替换公司名/excel';

const workSheets = xlsx.parse(sourceFilePath);
const nameMaps = {}

workSheets[0].data.forEach((line, index) => {
    if (index > 1 && line[3]) {
        if ((line[3].trim()) !== line[2].trim()) {
            nameMaps[(line[3].trim())] = line[2].trim()
        }
    }
})

const nameKeys = Object.keys(nameMaps) 

const files = fs.readdirSync(folderPath)

const logs = [
    {
        name: '没找到的医院',
        data: []
    }
]

files.forEach((fileName) => {
    if (fileName.includes('.xlsx')) {
        const filPath = `${folderPath}\/${fileName}`
        const sheets = xlsx.parse(filPath);
        sheets[0].data.forEach((_, index) => {
            if (index > 0) {
                const name = (sheets[0].data[index][2] || '').trim()
                if (name) {
                    nameKeys.forEach((key) => {
                        if(name.includes(key) || key.includes(name) || nameMaps[key].includes(name)) {
                            sheets[0].data[index][2] = nameMaps[key]
                            return false;
                        }
                    })
    
                    if (name === sheets[0].data[index][2]) {
                        logs[0].data.push([
                            fileName,
                            `${index} 行`,
                            name,
                            `${fileName} ${index}行: ${name} 未找到对应全称`
                        ])
                    }
                }
                sheets[0].data[index][7] = "微博"
            }
        })

        const names = filPath.split(".xlsx");
        names.pop();
        fs.writeFileSync(`${names.join('')}-replace-name.xlsx`, xlsx.build(sheets), "binary");
    }
})

if (logs[0].data.length > 0) {
    fs.writeFileSync(`${folderPath}/no_found_names.xlsx`,  xlsx.build(logs), { encoding: 'buffer' })
}
