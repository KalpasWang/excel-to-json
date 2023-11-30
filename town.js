import * as XLSX from "xlsx";

/* load 'fs' for readFile and writeFile support */
import * as fs from "fs";
XLSX.set_fs(fs);

/* load 'stream' for stream support */
import { Readable } from "stream";
XLSX.stream.set_readable(Readable);

/* load the codepage support library for extended support with older formats  */
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";
XLSX.set_cptable(cpexcel);

const filenames = [
  "宜蘭縣.xls",
  "花蓮縣.xls",
  "金門縣.xls",
  "南投縣.xls",
  "屏東縣.xls",
  "苗栗縣.xls",
  "桃園市.xls",
  "高雄市.xls",
  "基隆市.xls",
  "連江縣.xls",
  "雲林縣.xls",
  "新北市.xls",
  "新竹市.xls",
  "新竹縣.xls",
  "嘉義市.xls",
  "嘉義縣.xls",
  "彰化縣.xls",
  "臺中市.xls",
  "臺北市.xls",
  "臺東縣.xls",
  "臺南市.xls",
  "澎湖縣.xls",
];

const townsJson = [];
const template = {
  countyName: "台北市",
  townName: "松山區",
  candidate1: 5436,
  candidate2: 55918,
  candidate3: 64207,
  validVotes: 125561,
  invalidVotes: 1762,
  totalVotes: 127323,
  totalElectors: 164654,
  votingRate: 77.3276,
};

for (let filename of filenames) {
  const workbook = XLSX.readFile(filename);
  const sheet_name_list = workbook.SheetNames;
  const worksheet = workbook.Sheets[sheet_name_list[0]];
  let data = XLSX.utils.sheet_to_json(worksheet);
  data = data.slice(3);
  console.log(data);
}

// write townsJson to file
// fs.writeFileSync("towns.json", JSON.stringify(townsJson));
