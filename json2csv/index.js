const axios = require('axios')
const fs = require('fs')
const execSync = require('child_process').execSync
const json2csv = require('json2csv')
const colors = require('colors')
const { format, compareAsc } = require('date-fns')

const { curl } = JSON.parse(fs.readFileSync(`${__dirname}/setting.json`))

const str = curl.match(/page=(\S*)&/)
let rows = []

if (str && str[1]) {
    let page = parseInt(str[1])
    let request = curl
    let pageString = str[0]

    // PLAN A
    // while(true) {
    //     const response = JSON.parse(execSync(request))
    //     if (response.rows.length < 1) {
    //         break
    //     }
    //     rows = rows.concat(response.rows)
    //     page += 1
    //     const newPageString = `page=${page}&`
    //     request = request.replace(pageString, newPageString)
    //     pageString = newPageString
    // }

    // PLAN B
    const { total } = JSON.parse(execSync(curl))
    const totalStr = curl.match(/rows=(\S*)'/)
    request = curl.replace(totalStr[0], `rows=${total}'`)
    console.log(request.yellow)
    const response = JSON.parse(execSync(request))
    rows = rows.concat(response.rows)
    const parser = new json2csv.Parser({ fields: Object.keys(rows[0]) });
    const csv = parser.parse(rows);

    const fileName = `${require('os').homedir()}/Downloads/${format(new Date(), 'MM_DD_hh_mm_ss')}.csv`
    fs.writeFileSync(fileName, csv, { encoding: 'utf-8' })
    console.log(`输入以下命令打开文件:\n open ${fileName}`.red)
}