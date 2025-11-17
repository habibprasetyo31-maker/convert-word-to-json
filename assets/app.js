// ==== DRAG & DROP ====
const dropArea=document.getElementById("dropArea");
dropArea.ondragover=e=>{e.preventDefault(); dropArea.classList.add("bg-blue-100");};
dropArea.ondragleave=()=>dropArea.classList.remove("bg-blue-100");
dropArea.ondrop=async e=>{
 e.preventDefault(); dropArea.classList.remove("bg-blue-100");
 const file=e.dataTransfer.files[0];
 if(!file||!file.name.endsWith(".docx")) return alert("File bukan .docx");
 const buf=await file.arrayBuffer();
 const doc=await docx.Document.load(buf);
 let text="";
 doc.paragraphs.forEach(p=>text+=p.text+"\n");
 document.getElementById("input").value=text;
};

// ==== BACA FILE=== 
async function readDocx(){
 const file=document.getElementById("docxFile").files[0];
 if(!file) return alert("Pilih file");
 const buf=await file.arrayBuffer();
 const doc=await docx.Document.load(buf);
 let t=""; doc.paragraphs.forEach(p=>t+=p.text+"\n");
 document.getElementById("input").value=t;
}

// ==== CONVERT ====
function convert(){
 const text=document.getElementById("input").value;
 const blocks=text.split(/\n\s*\n/);
 let qs=[],id=1;

 blocks.forEach(b=>{
  const lines=b.trim().split("\n");
  if(lines.length<5) return;
  const q=lines[0].replace(/^\d+\.\s*/,"");
  const opts=lines.filter(l=>/^[A-Da-d]\./.test(l)).map(o=>o.replace(/^[A-Da-d]\./,"").trim());
  let correct="";
  const ans=lines.find(l=>/jawaban|kunci|answer/i.test(l));
  if(ans){
    const L=ans.match(/[A-D]/i)?.[0].toUpperCase();
    if(L){ correct=opts["ABCD".indexOf(L)] }
  }
  qs.push({id:id++, text:q, options:opts, correct:correct||opts[0]});
 });

 document.getElementById("output").textContent=JSON.stringify(qs,null,2);
}

// ==== DOWNLOAD JSON ====
function downloadJSON(){
 const data=document.getElementById("output").textContent;
 if(!data.trim()) return alert("Belum ada JSON");
 const blob=new Blob([data],{type:"application/json"});
 const url=URL.createObjectURL(blob);
 const a=document.createElement("a");
 a.href=url; a.download="questions.json"; a.click();
}

// ==== EXPORT EXCEL ====
function exportExcel(){
 const data=document.getElementById("output").textContent;
 if(!data.trim()) return alert("Belum ada data");
 const arr=JSON.parse(data);
 const ws=XLSX.utils.json_to_sheet(arr);
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,ws,"Questions");
 XLSX.writeFile(wb,"questions.xlsx");
}
