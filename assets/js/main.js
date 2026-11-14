const pages=document.querySelectorAll('.page');
function go(id){
  if(!document.getElementById(id)) id='p-home';
  pages.forEach(p=>p.classList.toggle('on',p.id===id));
  document.querySelectorAll('.nv a').forEach(a=>a.classList.toggle('on',a.getAttribute('href')==='#'+id));
  window.scrollTo(0,0);
}
document.querySelectorAll('a[href^="#p-"]').forEach(a=>{
  a.addEventListener('click',e=>{e.preventDefault();const id=a.getAttribute('href').slice(1);history.replaceState(null,'','#'+id);go(id);});
});
const start=(location.hash||'#p-home').slice(1);
history.replaceState(null,'',location.pathname+location.search);
go(document.getElementById(start)?start:'p-home');

/* ---------- 资料库 ---------- */


/* ---------- 搜索 ---------- */
const q=document.getElementById('q'), qres=document.getElementById('qres'), qbtn=document.getElementById('qbtn');
function norm(s){return (s||'').toLowerCase()}
function hit(d,k){
  const hay=norm([d.t,d.cat,d.sum,d.kw,d.place.b,d.place.dept,d.mat.join(' ')].join(' '));
  return k.split(/\s+/).filter(Boolean).every(w=>hay.indexOf(norm(w))>-1);
}
function mark(s,k){ if(!k) return s; const w=k.trim().split(/\s+/)[0]; if(!w) return s;
  try{ return s.replace(new RegExp('('+w.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi'),'<em>$1</em>') }catch(e){ return s } }
function search(){
  const k=q.value.trim();
  if(!k){ qres.classList.remove('on'); qres.innerHTML=''; return; }
  const kk=k.toLowerCase().split(/\s+/).filter(Boolean);
  const r=DOCS.map(d=>{let s=0;kk.forEach(w=>{
    if(d.t.toLowerCase().indexOf(w)>-1) s+=10;
    if((d.cat+' '+d.sum).toLowerCase().indexOf(w)>-1) s+=4;
    if(d.kw.toLowerCase().indexOf(w)>-1) s+=3;
    if(d.mat.join(' ').toLowerCase().indexOf(w)>-1) s+=1;});
    return {d:d,s:s};}).filter(x=>x.s>0).sort((a,b)=>b.s-a.s).map(x=>x.d);
  qres.classList.add('on');
  if(!r.length){ qres.innerHTML='<div class="qnone">没有找到「'+k+'」，换个说法试试：学生证、助学金、报修、请假</div>'; return; }
  qres.innerHTML='<div style="font-size:12px;color:#8b98a8;margin:0 0 8px 6px">找到 '+r.length+' 条</div><div class="rc">'+r.map(d=>
    '<div class="ri" data-id="'+d.id+'"><span class="cat">'+d.cat+'</span><div style="min-width:0"><div class="rt">'+mark(d.t,k)+'</div><div class="rs">'+mark(d.sum,k)+'</div></div></div>'
  ).join('')+'</div>';
}
q.addEventListener('input',search);
q.addEventListener('keydown',e=>{if(e.key==='Enter'){const f=qres.querySelector('.ri'); if(f) openDoc(f.dataset.id);}});
qbtn.addEventListener('click',search);
qres.addEventListener('click',e=>{const it=e.target.closest('.ri'); if(it) openDoc(it.dataset.id);});
document.querySelectorAll('.kw').forEach(a=>a.addEventListener('click',e=>{
  e.preventDefault(); q.value=a.textContent.trim(); search(); q.focus();
  document.querySelector('.sb').scrollIntoView({block:'center'});
}));

/* ---------- 详情弹层 ---------- */
const mask=document.getElementById('mask'), mbody=document.getElementById('mbody');
function openDoc(id){
  const d=DOCS.find(x=>x.id===id); if(!d) return;
  mbody.innerHTML =
   '<span class="x" id="mx">×</span>'+
   '<div class="mc">'+d.cat+'</div><h2>'+d.t+'</h2><div class="ms">'+d.sum+'</div>'+
   '<div class="blk"><h4>适用对象</h4><div style="font-size:13.5px;color:#3a4b60;line-height:1.9">'+d.who+'</div></div>'+
   '<div class="blk"><h4>需要带的材料</h4><ul>'+d.mat.map(m=>'<li>'+m+'</li>').join('')+'</ul></div>'+
   '<div class="blk"><h4>去哪办</h4><div class="place"><b>地点</b><span>'+d.place.b+' '+d.place.f+' '+d.place.room+'</span><b>部门</b><span>'+d.place.dept+'</span><b>联系方式</b><span>'+(d.tel||'—')+'</span><b>时间</b><span>'+d.time+'</span></div></div>'+
   '<div class="blk"><h4>办理步骤</h4><ol>'+d.steps.map(s=>'<li>'+s+'</li>').join('')+'</ol></div>'+
   (d.note?'<div class="blk"><h4>注意事项</h4><div style="font-size:13.5px;color:#3a4b60;line-height:1.9">'+d.note+'</div></div>':'')+
   '<div class="gen"><button class="btn" id="mgen">一键生成办理单</button><span>生成一份 .doc 文件，可保存、可打印，带去办事窗口</span></div>';
  mask.classList.add('on');
  document.getElementById('mx').onclick=closeDoc;
  document.getElementById('mgen').onclick=()=>makeDoc(d);
}
function closeDoc(){mask.classList.remove('on');}
mask.addEventListener('click',e=>{if(e.target===mask) closeDoc();});
document.addEventListener('keydown',e=>{if(e.key==='Escape') closeDoc();});

/* ---------- 生成文件 ---------- */
function makeDoc(d){
  const now=new Date();
  const pad=n=>String(n).padStart(2,'0');
  const stamp=now.getFullYear()+'-'+pad(now.getMonth()+1)+'-'+pad(now.getDate())+' '+pad(now.getHours())+':'+pad(now.getMinutes());
  const html='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>'+d.t+' 办理单</title>'
   +'<style>body{font-family:"宋体",SimSun;font-size:11pt;line-height:1.75;color:#000;margin:2.2cm 2.4cm}h1{font-family:"黑体",SimHei;font-size:16pt;text-align:center;margin:0 0 6pt}p.sub{text-align:center;color:#555;font-size:9.5pt;margin:0 0 16pt}h2{font-family:"黑体",SimHei;font-size:12pt;margin:16pt 0 6pt;border-left:4px solid #0d4a7a;padding-left:8pt}table{border-collapse:collapse;width:100%}td{border:1px solid #999;padding:6pt;font-size:10.5pt}td.k{background:#f2f5f8;width:22%;font-family:"黑体",SimHei}ul{margin:0;padding-left:18pt}li{margin:3pt 0}ol{margin:0;padding-left:18pt}.ft{margin-top:22pt;border-top:1px solid #bbb;padding-top:8pt;font-size:9pt;color:#666}</style></head><body>'
   +'<h1>'+d.t+' · 办理单</h1><p class="sub">江苏省张家港中等专业学校　'+d.cat+'　生成时间：'+stamp+'</p>'
   +'<h2>一、适用对象</h2><p>'+d.who+'</p>'
   +'<h2>二、需要带的材料</h2><ul>'+d.mat.map(m=>'<li>□ '+m+'</li>').join('')+'</ul>'
   +'<h2>三、办理地点</h2><table><tr><td class="k">地点</td><td>'+d.place.b+' '+d.place.f+' '+d.place.room+'</td></tr><tr><td class="k">承办部门</td><td>'+d.place.dept+'</td></tr><tr><td class="k">联系方式</td><td>'+(d.tel||'—')+'</td></tr><tr><td class="k">办理时间</td><td>'+d.time+'</td></tr></table>'
   +'<h2>四、办理步骤</h2><ol>'+d.steps.map(s=>'<li>'+s+'</li>').join('')+'</ol>'
   +(d.note?'<h2>五、注意事项</h2><p>'+d.note+'</p>':'')
   +'<div class="ft">本办理单由江苏省张家港中等专业学校校园门户生成，内容依据校内办事流程整理。<br>部分房间号与时间为演示数据，实际以现场标识和窗口公告为准。</div></body></html>';
  try{
    const blob=new Blob(['\ufeff'+html],{type:'application/msword;charset=utf-8'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download=d.t.replace(/[\/\\:*?"<>|（）()]/g,'')+'-办理单.doc';
    document.body.appendChild(a); a.click();
    setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},1200);
  }catch(e){ alert('生成失败：'+e.message); }
}

/* ---------- 公告与新闻正文 ---------- */

try{const s=localStorage.getItem('zjg_arts_v1');if(s){const d=JSON.parse(s);if(Array.isArray(d)&&d.length){ARTS.length=0;d.forEach(x=>ARTS.push(x));}}}catch(e){}
const ART_MAP={}; ARTS.forEach(a=>ART_MAP[a.id]=a);

function artHref(a){ return 'https://zjgzzx.zjgedu.cn/xwzx/'+(a.g==='ann'?'xygg':'xyxw')+'/content_'+a.cid+'.html'; }
function renderList(el,kind){
  if(!el) return;
  el.innerHTML=ARTS.filter(a=>a.g===kind).map(a=>
    '<li><span class="pill">'+a.cat+'</span><a class="p" href="#" data-a="'+a.id+'">'+a.t+'</a><span class="d">'+a.d+'</span></li>'
  ).join('');
}
renderList(document.getElementById('list-ann'),'ann');
const _ca=document.getElementById('cnt-ann'); if(_ca) _ca.textContent=ARTS.filter(a=>a.g==='ann').length;
renderList(document.getElementById('list-news'),'news');
renderList(document.getElementById('list-dept'),'dept');

let artBack='#p-notice';
function openArt(id){
  const a=ART_MAP[id]; if(!a) return;
  document.getElementById('art-title').textContent=a.t;
  const gname=a.g==='ann'?'校园公告':(a.g==='news'?'校园新闻':'部门新闻');
  document.getElementById('art-meta').innerHTML='2026-'+a.d+' · '+gname+' · '+a.dept+(a.cid?' · 来源：<a href="'+artHref(a)+'" target="_blank" style="color:#0d4a7a;text-decoration:none;border-bottom:1px solid #cbd5e1">官网原文</a>':'');
  document.getElementById('art-body').innerHTML=a.body.map(p=>'<p>'+p+'</p>').join('');
  const gen=document.getElementById('art-gen');
  if(a.gen){
    gen.innerHTML='<div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap">'
      +'<button class="btn" id="art-genbtn" style="background:#0d4a7a;color:#fff;border:0;cursor:pointer;font-family:inherit;font-size:14px;padding:12px 22px;border-radius:9px">生成'+a.gen.title+'</button>'
      +'<span style="font-size:12.5px;color:var(--sub)">下载一份 .doc，可打印、可填写</span></div>';
    document.getElementById('art-genbtn').onclick=()=>makeArtDoc(a);
  } else {
    gen.innerHTML='<div style="font-size:12.5px;color:var(--sub)">本条不涉及学生办理事项，无需生成文件。</div>';
  }
  const back=document.getElementById('art-back');
  back.href=artBack;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('on'));
  document.getElementById('p-article').classList.add('on');
  document.querySelectorAll('.nv a').forEach(x=>x.classList.remove('on'));
  window.scrollTo(0,0);
}
document.addEventListener('click',e=>{
  const t=e.target.closest('a[data-a]');
  if(t){ e.preventDefault(); openArt(t.dataset.a); }
});
/* 首页上的新闻、公告也接上 */
['.ncard a.t','.nlist a','.glist a'].forEach(sel=>{
  document.querySelectorAll(sel).forEach(a=>{
    const hit=ARTS.find(x=>x.t===a.textContent.trim());
    if(hit){ a.addEventListener('click',e=>{e.preventDefault(); artBack='#p-home'; openArt(hit.id);}); }
  });
});

function makeArtDoc(a){
  const g=a.gen, now=new Date(), pad=n=>String(n).padStart(2,'0');
  const stamp=now.getFullYear()+'-'+pad(now.getMonth()+1)+'-'+pad(now.getDate())+' '+pad(now.getHours())+':'+pad(now.getMinutes());
  let body='';
  if(g.kind==='signup'){
    const FIELDS=g.fields||g.items||['姓名','班级','联系电话','报名项目','备注'];
    body='<h2>一、填写事项</h2><table>'+FIELDS.map(f=>'<tr><td class="k">'+f+'</td><td>&nbsp;</td></tr>').join('')+'</table>'
        +'<p style="margin-top:14pt;font-size:10pt;color:#555">请用黑色签字笔填写，或打印后手写；与身份证信息保持一致。</p>';
  } else {
    body='<h2>一、材料清单</h2><ul>'+(g.items||[]).map(i=>'<li>&#9633; '+i+'</li>').join('')+'</ul>'
        +'<p style="margin-top:14pt;font-size:10pt;color:#555">按清单顺序整理，用文件夹装订，交至承办部门。</p>';
  }
  const html='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>'+g.title+'</title>'
   +'<style>body{font-family:"宋体",SimSun;font-size:11pt;line-height:1.75;margin:2.2cm 2.4cm}h1{font-family:"黑体",SimHei;font-size:16pt;text-align:center;margin:0 0 6pt}p.sub{text-align:center;color:#555;font-size:9.5pt;margin:0 0 16pt}h2{font-family:"黑体",SimHei;font-size:12pt;margin:16pt 0 6pt;border-left:4px solid #0d4a7a;padding-left:8pt}table{border-collapse:collapse;width:100%}td{border:1px solid #999;padding:8pt;font-size:10.5pt}td.k{background:#f2f5f8;width:32%;font-family:"黑体",SimHei}ul{margin:0;padding-left:18pt}li{margin:4pt 0}.ft{margin-top:24pt;border-top:1px solid #bbb;padding-top:8pt;font-size:9pt;color:#666}</style></head><body>'
   +'<h1>'+g.title+'</h1><p class="sub">江苏省张家港中等专业学校　生成时间：'+stamp+'</p>'
   +'<h2>依据</h2><p>'+a.t+'（2026-'+a.d+'）</p>'
   +body
   +'<div class="ft">承办部门：'+a.dept+'　｜　本表依据学校官网公告整理，具体以公告原文为准。</div></body></html>';
  try{
    const blob=new Blob(['\ufeff'+html],{type:'application/msword;charset=utf-8'});
    const el=document.createElement('a');
    el.href=URL.createObjectURL(blob);
    el.download=a.sn+'-'+g.title+'.doc';
    document.body.appendChild(el); el.click();
    setTimeout(()=>{URL.revokeObjectURL(el.href); el.remove();},1200);
  }catch(e){ alert('生成失败：'+e.message); }
}
