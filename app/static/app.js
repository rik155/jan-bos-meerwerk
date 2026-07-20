const form=document.getElementById('form');
const filesInput=document.getElementById('fotos');
const preview=document.getElementById('preview');
const canvas=document.getElementById('signature');
const ctx=canvas.getContext('2d');
const msg=document.getElementById('message');
const submit=document.getElementById('submit');
let selected=[]; let drawing=false; let signed=false;

document.querySelector('[name="datum"]').valueAsDate=new Date();
function renderFiles(){
  preview.innerHTML='';
  selected.forEach((file,i)=>{const box=document.createElement('div');box.className='thumb';const img=document.createElement('img');img.src=URL.createObjectURL(file);const b=document.createElement('button');b.type='button';b.textContent='×';b.onclick=()=>{selected.splice(i,1);renderFiles()};box.append(img,b);preview.append(box)});
}
filesInput.addEventListener('change',()=>{selected=[...selected,...filesInput.files].slice(0,12);filesInput.value='';renderFiles()});

function point(e){const r=canvas.getBoundingClientRect();const p=e.touches?e.touches[0]:e;return{x:(p.clientX-r.left)*(canvas.width/r.width),y:(p.clientY-r.top)*(canvas.height/r.height)}}
function start(e){e.preventDefault();drawing=true;const p=point(e);ctx.beginPath();ctx.moveTo(p.x,p.y)}
function move(e){if(!drawing)return;e.preventDefault();const p=point(e);ctx.lineWidth=5;ctx.lineCap='round';ctx.strokeStyle='#111';ctx.lineTo(p.x,p.y);ctx.stroke();signed=true}
function stop(){drawing=false}
canvas.addEventListener('pointerdown',start);canvas.addEventListener('pointermove',move);window.addEventListener('pointerup',stop);
document.getElementById('clear').onclick=()=>{ctx.clearRect(0,0,canvas.width,canvas.height);signed=false};

form.addEventListener('submit',async e=>{
  e.preventDefault(); msg.hidden=true;
  if(!signed){show('Zet eerst een handtekening.',true);return}
  const fd=new FormData(form);
  const data={
    opdrachtgever:fd.get('opdrachtgever'),opdrachtnummer:fd.get('opdrachtnummer')||'',object:fd.get('object'),datum:fd.get('datum'),medewerker:fd.get('medewerker'),
    omschrijving:fd.get('omschrijving'),reden:fd.get('reden')||'',uren:Number(fd.get('uren')||0),materialen:fd.get('materialen')||'',kostenindicatie:fd.get('kostenindicatie')||'',handtekening:canvas.toDataURL('image/png')
  };
  const body=new FormData();body.append('data_json',JSON.stringify(data));selected.forEach(f=>body.append('fotos',f,f.name));
  submit.disabled=true;submit.textContent='Bezig met verzenden...';
  try{const res=await fetch('/api/meerwerk',{method:'POST',body});const out=await res.json();if(!res.ok)throw new Error(out.detail||'Verzenden mislukt');show(`Meerwerkrapport verzonden. Bestand: ${out.bestandsnaam}`,false);form.reset();document.querySelector('[name="datum"]').valueAsDate=new Date();selected=[];renderFiles();ctx.clearRect(0,0,canvas.width,canvas.height);signed=false;window.scrollTo({top:0,behavior:'smooth'})}catch(err){show(err.message,true)}finally{submit.disabled=false;submit.textContent='Meerwerk verzenden'}
});
function show(text,error){msg.textContent=text;msg.className='message'+(error?' error':'');msg.hidden=false;msg.scrollIntoView({behavior:'smooth',block:'center'})}
if('serviceWorker' in navigator)navigator.serviceWorker.register('/static/sw.js');
