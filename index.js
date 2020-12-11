let express = require("express");
let app = express();
let url = require("url");
let ejs = require('ejs');
app.engine('ejs',ejs.renderFile);
let bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({extended:false}));
require('dotenv').config();
const {GoogleSpreadsheet} = require('google-spreadsheet');

app.get("/",(req,res,next)=>{
    res.render('index.ejs',{
        title:"活動出欠",
    })
})
app.post("/",(req,res,next)=>{
    let url_sheet = url.parse(req.body.url,true);
    let spreadsheet_id = url_sheet.pathname.split("/")[3];
    let worksheet_id = url_sheet.hash.split('=')[1];

    async function loadSpreadsheet(){
        const doc = new GoogleSpreadsheet(spreadsheet_id);
        const credentials = require('./credentials.json');
        await doc.useServiceAccountAuth(credentials);
        await doc.loadInfo();
      
        const attendanceSheet = await doc.sheetsById[worksheet_id];
        const attendanceRows = await attendanceSheet.getRows();
        let attendant_row = new Array();
        for (let i=0;i<attendanceRows.length;i++){
          if (attendanceRows[i].参加不参加をご入力ください==="参加"){
            attendant_row.push(attendanceRows[i]);
          }
        }
        let attendant = new Array();
        for (let i=0;i<attendant_row.length;i++){
          attendant.push(attendant_row[i]._rawData.slice(1,2)[0]);
        }
      
        return attendant;
    }
    
    async function jikkou(){
        let attendants = new Array();
        await loadSpreadsheet()
        .then(attendant=>{
            for (let i of attendant){
                attendants.push(i)
            }
            attendants = attendants.filter(function (x, i, self) {
                return self.indexOf(x) === i;
              });
        })
        .catch(err=>console.error(err));
        
        res.render('index2.ejs',{
            title:'活動出欠',
            url:req.body.url,
            attendant:attendants,
        })
    }
    jikkou();
})

app.listen(3000,()=>{
    console.log('Start server port:3000')
})