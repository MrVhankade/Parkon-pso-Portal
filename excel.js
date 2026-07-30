
let data=[];
async function loadExcel(){
 const r=await fetch("HTP_Height_Chart.xlsx")
 const b=await r.arrayBuffer();
 const wb=XLSX.read(b);
 const ws=wb.Sheets[wb.SheetNames[0]];
 data=XLSX.utils.sheet_to_json(ws);
}
loadExcel();

function searchHeight(){
 const h=Number(document.getElementById('height').value);
 const row=data.find(x=>Number(x['AVAILABLE CLEAR HEIGHT (H)'])===h);
 const out=document.getElementById('result');
 if(!row){out.innerHTML='<b>No matching height found.</b>';return;}
 out.innerHTML=`
 <h3>Result</h3>
 <p><b>Available Height:</b> ${row['AVAILABLE CLEAR HEIGHT (H)']} mm</p>
 <p><b>System Type:</b> ${row['SYSTEM TYPE']}</p>
 <p><b>GF Car Height:</b> ${row['ALLOWABLE CAR HEIGHT ON GROUND FLOOR GF CAR']} mm</p>
 <p><b>1F Car Height:</b> ${row['ALLOWABLE CAR HEIGHT ON GROUND FLOOR 1F CAR']} mm</p>`;
}
