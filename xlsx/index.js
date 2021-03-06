const fs = require("fs");
const xlsx = require("node-xlsx").default;

const splitExcel = path => {
  const workSheets = xlsx.parse(path);

  const splitSheets = [
    {
      name: workSheets[0].name,
      data: []
    }
  ];

  workSheets[0].data.forEach((line, index) => {
    if (index === 0) {
      splitSheets[0].data.push([
        "客户姓名",
        "电话",
        "派单医院",
        "重复状态",
        "跟进状态",
        "派单时间",
        "跟进时间",
        "来源渠道",
        "咨询项目",
        "咨询地区",
        "咨询情况",
        "所属人",
        "推广人",
        "录单人",
        "录单时间"
      ]);
    } else {
      if (line[2] && ~line[2].indexOf("\n")) {
        line[2].split("\n").forEach(text => {
          if (text !== "") {
            text = text.replace(/:/g, '&')
            text = text.replace(/；/g, ';')
            text = text.replace(/：/g, ':')
            const splits = text.split(";");
            line[2] = splits[0].trim();
            splitSheets[0].data.push([...line]);
            splitSheets[0].data[splitSheets[0].data.length - 1].splice(
              3,
              0,
              splits[1].trim().split(":")[1],
              splits[2].trim().split(":")[1],
              splits[3].trim().split(":")[1].replace(/&/g, ':'),
              splits[4].trim().split(":")[1].replace(/&/g, ':'),
            );
          }
        });
      } else {
        splitSheets[0].data.push([...line]);
        splitSheets[0].data[splitSheets[0].data.length - 1].splice(
          3,
          0,
          "",
          "",
          "",
          ""
        );
      }
    }
  });

  const buffer = xlsx.build(splitSheets);
  const names = path.split(".xlsx");
  names.pop();
  fs.writeFileSync(`${names.join('')}-split.xlsx`, buffer, "binary");
};

// const folderPath = process.argv.pop();
const folderPath = '/Users/huanminghu/Library/Containers/com.tencent.xinWeChat/Data/Library/Application\ Support/com.tencent.xinWeChat/2.0b4.0.9/76baac86150b63d88137e22b38c746f8/Message/MessageTemp/b76931d9ba7ea39622c94275c86872c1/File/表格'

const files = fs.readdirSync(folderPath)

files.forEach((fileName) => {
    if (fileName.includes('.xlsx')) {
        splitExcel(`${folderPath}\/${fileName}`)
    }
})