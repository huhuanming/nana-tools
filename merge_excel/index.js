const fs = require("fs")
const path = require("path")
const xlsx = require("node-xlsx").default

const dirPath = '/Users/huanminghu/Downloads/branch_excel'
const files = fs.readdirSync(dirPath)


const workSheets = []
files.forEach((fileName) => {
    if (fileName.indexOf('.') === 0) {
        return
    }
    const filePath = path.join(dirPath, fileName)
    const sourceSheets = xlsx.parse(filePath)
    sourceSheets.forEach((sourceSheet, index) => {
        const sheet = workSheets[index]
        const lastCommentLine = sourceSheet.data.findIndex((line) => line[0] && (~line[0].indexOf('统计人：') || ~line[0].indexOf('填写说明：')))
        const lastBlankLine = sourceSheet.data.findIndex((line) => line.join('') === '')
        const lastLine = lastCommentLine > lastBlankLine ? lastBlankLine : lastCommentLine
        if (!sheet) {
            workSheets[index] = sourceSheet
            workSheets[index].data = workSheets[index].data.slice(0, lastLine)
        } else {
            switch(index) {
                case 0:
                    sheet.data = sheet.data.concat(sourceSheet.data.slice(3, lastLine))
                    break
                default:
                    sheet.data = sheet.data.concat(sourceSheet.data.slice(2, lastLine))
                    break
            }
        }
    })
})


fs.writeFileSync('境内分行机构新型冠状病毒肺炎情况统计表.xlsx', xlsx.build(workSheets), { encoding: 'buffer' })