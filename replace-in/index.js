const fs = require("fs");
const xlsx = require("node-xlsx").default;

// const sourceFilePath = process.argv.pop();
// const folderPath = process.argv.pop();

const sourceFilePath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/推月聚流-关联医院.xlsx';
const folderPath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/replace-name-15-47';

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

const checkName = (name, key, value) => {
    const editName = name.replace('整形', '')
    if (editName.includes(key) || key.includes(editName) || value.includes(editName)) {
        return true
    }
    return name.includes(key) || key.includes(name) || value.includes(name)
}

files.forEach((fileName) => {
    if (fileName.includes('.xlsx')) {
        const filPath = `${folderPath}\/${fileName}`
        const sheets = xlsx.parse(filPath);
        sheets[0].data.forEach((_, index) => {
            if (index > 0) {
                const name = (sheets[0].data[index][2] || '').trim()
                if (name) {
                    let matchKey = ''
                    nameKeys.forEach((key) => {
                        if(checkName(name, key, nameMaps[key])) {
                            matchKey = nameMaps[key]
                            sheets[0].data[index][2] = nameMaps[key]
                            return false;
                        }
                    })
    
                    if (name !==  matchKey && name === sheets[0].data[index][2].trim()) {
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
