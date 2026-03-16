
(function(){
  const byId = (id)=>document.getElementById(id);

  function mock(){
    return {
      summary:{secureScore:72,compliancePercent:78,userCount:1250,assessmentDate:new Date().toISOString(),passCount:156,warningCount:18,failCount:24},
      categories:[
        {name:'Identity & Access',passCount:42,warningCount:2,failCount:3},
        {name:'Data Protection',passCount:38,warningCount:3,failCount:5},
        {name:'Threat Protection',passCount:35,warningCount:4,failCount:8}
      ],
      recommendations:[
        {title:'Enable MFA for all admins'},
        {title:'Add Conditional Access baseline'},
        {title:'Configure DLP policies'}
      ],
      history:[
        {date:new Date(Date.now()-7*86400000).toISOString(),secureScore:70,passCount:150,failCount:26},
        {date:new Date(Date.now()-14*86400000).toISOString(),secureScore:68,passCount:148,failCount:28}
      ]
    };
  }

  async function loadData(){
    try{
      const r=await fetch('_snapshots/M365-Complete-Baseline-latest.summary.json');
      if(!r.ok) throw new Error('http '+r.status);
      return await r.json();
    }catch(_){
      return mock();
    }
  }

  function render(data){
    const s=data.summary||{};
    byId('secureScore').textContent=s.secureScore ?? '-';
    byId('compliancePercent').textContent=(s.compliancePercent ?? '-') + '%';
    byId('userCount').textContent=(s.userCount ?? '-').toLocaleString ? s.userCount.toLocaleString('nl-NL') : s.userCount;
    byId('assessmentDate').textContent=new Date(s.assessmentDate || Date.now()).toLocaleString('nl-NL');
    byId('passCount').textContent=s.passCount ?? '-';
    byId('warningCount').textContent=s.warningCount ?? '-';
    byId('failCount').textContent=s.failCount ?? '-';

    const ct=byId('categoryTable');
    const cats=data.categories||[];
    ct.innerHTML=cats.length?cats.map(c=>`<tr><td>${c.name}</td><td>${c.passCount||0}</td><td>${c.warningCount||0}</td><td>${c.failCount||0}</td></tr>`).join(''):'<tr><td colspan="4">Geen data</td></tr>';

    const rl=byId('recommendationsList');
    const recs=data.recommendations||[];
    rl.innerHTML=recs.length?recs.slice(0,5).map((r,i)=>`<div style="padding:10px;border-left:3px solid #FF6600;background:#fff7f0;margin-bottom:8px;border-radius:6px">${i+1}. ${r.title||''}</div>`).join(''):'Geen aanbevelingen';

    const ht=byId('historyTable');
    const hist=data.history||[];
    ht.innerHTML=hist.length?hist.map(h=>`<tr><td>${new Date(h.date).toLocaleDateString('nl-NL')}</td><td>${h.secureScore||0}</td><td>${h.passCount||0}</td><td>${h.failCount||0}</td></tr>`).join(''):'<tr><td colspan="4">Geen historie</td></tr>';
  }

  function wire(data){
    document.querySelectorAll('[data-section]').forEach(el=>{
      el.addEventListener('click',(e)=>{
        e.preventDefault();
        document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
        el.classList.add('active');
        const section=el.getAttribute('data-section');
        byId('overviewSection').style.display=section==='overview'?'block':'none';
        byId('historySection').style.display=section==='history'?'block':'none';
      });
    });
    byId('refreshBtn').addEventListener('click',()=>location.reload());
    byId('downloadBtn').addEventListener('click',()=>{
      const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
      const url=URL.createObjectURL(blob);
      const a=document.createElement('a');
      a.href=url; a.download='m365-baseline.json'; a.click();
      URL.revokeObjectURL(url);
    });
  }

  document.addEventListener('DOMContentLoaded', async ()=>{
    const data=await loadData();
    render(data);
    wire(data);
  });
})();
