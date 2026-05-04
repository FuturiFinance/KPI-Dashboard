//# sourceURL=app.js
/* app.js — React UMD + Chart.js 2.x + PapaParse (Netlify static) */
(function () {
  'use strict';
  var e = React.createElement;

  /* ------------------ helpers ------------------ */
  function fmtCurrency(n){return n==null||isNaN(n)?'—':new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0}).format(n);}
  function fmtPercent(n){return n==null||isNaN(n)?'—':(n*100).toFixed(1)+'%';}
  function fmtCompact(n){
    if(n==null||isNaN(n))return'—';
    var a=Math.abs(n);
    if(a>=1e9)return'$'+(n/1e9).toFixed(1)+'B';
    if(a>=1e6)return'$'+(n/1e6).toFixed(1)+'M';
    if(a>=1e3)return'$'+(n/1e3).toFixed(1)+'K';
    return'$'+String(Math.round(n));
  }
  function fmtCompactMoney(n){
    if(n==null||isNaN(n))return'—';
    var a=Math.abs(n);
    if(a>=1e6)return'$'+(n/1e6).toFixed(1)+'M';
    if(a>=1e3)return'$'+(n/1e3).toFixed(1)+'K';
    return'$'+String(Math.round(n));
  }
  function fmtGIThousands(n) {
    if(n == null || isNaN(n)) return '—';
    return new Intl.NumberFormat('en-US').format(Math.round(n));
  }
  function isPctName(n){return /%|Margin|Rate|Growth|Rule of 40|Retention/i.test(n);}
  function safeDiv(a,b){return a==null||b==null||b===0?null:a/b;}
  function deltaPct(curr, prev){ if(curr==null||prev==null||prev===0)return null; return (curr-prev)/prev; }
  function clsDelta(n,goodUp){ if(n==null) return ''; var pos=n>0, bad=(goodUp? !pos: pos); return bad? 'text-rose-400' : 'text-emerald-400'; }
  function prettyDelta(n){ return n==null?'—':((n*100).toFixed(1)+'%'); }

  /* ------------------ parsing ------------------ */
  function canonicalMetricName(name){
    if(!name)return name;
    var n=String(name).trim().replace(/\s+/g,' ');
    n=n.replace(/Retentuion/i,'Retention')
       .replace(/Gross\s*Margin%/i,'Gross Margin %')
       .replace(/Products\s*\(customers\)/i,'Products (customers)')
       .replace(/^Capex$/i,'CAPEX').replace(/^CapEx$/i,'CAPEX');
    return n;
  }
  function sanitizeNumber(raw,metricName){
    if(raw==null||raw==='')return null;
    var s=String(raw).trim(),neg=false;
    if(s[0]==='('&&s[s.length-1]===')'){neg=true;s=s.slice(1,-1);}
    s=s.replace(/[$€£, ]/g,'');
    if(/%$/.test(s)){s=s.replace('%','');var v=parseFloat(s);if(isNaN(v))return null;v=v/100;return neg?-v:v;}
    if(isPctName(metricName)){var v2=parseFloat(s);if(isNaN(v2))return null;if(Math.abs(v2)>1.5)v2=v2/100;return neg?-v2:v2;}
    var v3=parseFloat(s);if(isNaN(v3))return null;return neg?-v3:v3;
  }
  function parseMonth(v){
    if(!v)return null;
    var s=String(v).trim(),m;
    m=/^(\d{4})-(\d{1,2})$/.exec(s);
    if(m)return new Date(+m[1],+m[2]-1,1);
    
    m=/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.exec(s);
    if(m){
      var y=+m[3];
      if(y<100)y+=2000;
      return new Date(y,+m[1]-1,1);
    }
    
    m=/^([A-Za-z]{3,})[ \/-](\d{2,4})$/.exec(s);
    if(m){
      var mo=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[1].slice(0,3).toLowerCase());
      if(mo>=0){
        var yy=+m[2];
        if(yy<100)yy+=2000;
        return new Date(yy,mo,1);
      }
    }
    
    var d=new Date(s);
    return isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),1);
  }
  function sameMonth(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth();}
  function monthKey(d){return d?d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2):'';}
  function firstOfMonth(ym){var p=ym.split('-');return new Date(+p[0],+p[1]-1,1);}
  function addMonths(d,n){return new Date(d.getFullYear(),d.getMonth()+n,1);}
  function monthShortLabel(v){
    var d=v;
    if(!(v instanceof Date)) d=parseMonth(v);
    return d?(d.getMonth()+1)+'/'+String(d.getFullYear()).slice(-2):String(v);
  }
  function monthLongLabel(d){
    if(!d) return '';
    var months=['January','February','March','April','May','June','July','August','September','October','November','December'];
    return months[d.getMonth()]+' '+d.getFullYear();
  }

  // Parse money values that may have formatting
  function parseMoney(val){
    if(val==null || val==='') return 0;
    if(typeof val === 'number') return val;
    var s = String(val).replace(/[$,"\s]/g,'');
    if(s === '#DIV/0!' || s === 'n/a' || s === '') return 0;
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  /* ------------------ chart utils ------------------ */
  if(window.Chart&&Chart.defaults&&Chart.defaults.global){
    Chart.defaults.global.defaultFontColor='#9fb0c7';
    Chart.defaults.global.defaultFontFamily='system-ui,-apple-system,Segoe UI,Roboto,sans-serif';
    Chart.defaults.global.animation.duration=0;
  }
  if(window.Chart&&Chart.plugins){
    Chart.plugins.register({
      afterUpdate:function(chart){
        var s=chart.scales||{};
        Object.keys(s).forEach(function(k){
          var a=s[k];
          if(!a||!a.ticks||!a.isHorizontal||!a.isHorizontal())return;
          var labels=(chart.data&&chart.data.labels)||[];
          if(!labels.length||!a.ticks.length)return;
          a.ticks[0]=labels[0];
          a.ticks[a.ticks.length-1]=labels[labels.length-1];
        });
      }
    });
  }
  function gridLines(){
    return {
      color:'rgba(148,163,184,.15)',
      zeroLineColor:'rgba(148,163,184,.25)'
    };
  }
  function destroyOnCanvas(canvasId){
    if(!window.Chart||!Chart.instances)return;
    var insts=Chart.instances;
    if(Array.isArray(insts)){
      for (var i=insts.length-1;i>=0;i--){
        var it=insts[i];
        if(it&&it.chart&&it.chart.canvas&&it.chart.canvas.id===canvasId){
          try{it.destroy();}catch(e){}
        }
      }
    }else{
      for (var k in insts){
        var it2=insts[k];
        if(it2&&it2.chart&&it2.chart.canvas&&it2.chart.canvas.id===canvasId){
          try{it2.destroy();}catch(e){}
        }
      }
    }
  }
  function findSizedAncestor(el){
    var cur=el;
    while(cur&&cur!==document.body){
      if((cur.clientWidth||0)>0)return cur;
      cur=cur.parentElement;
    }
    return el.parentElement||el;
  }
  function createWhenReady(canvasId,build){
    function attempt(tries){
      var el=document.getElementById(canvasId);
      if(!el){
        if(tries<60)return setTimeout(function(){attempt(tries+1);},50);
        console.warn('[chart] canvas not found',canvasId);
        return;
      }
      var holder=findSizedAncestor(el);
      if(!holder.style.height) holder.style.height='320px';
      
      function make(){
        destroyOnCanvas(canvasId);
        el.width=holder.clientWidth||800;
        el.height=holder.clientHeight||320;
        var ctx=el.getContext('2d');
        requestAnimationFrame(function(){build(ctx);});
      }
      
      if((holder.clientWidth||0)>0){
        make();
      }else{
        var ro=new ResizeObserver(function(r){
          if(!r.length)return;
          if(r[0].contentRect.width>0){
            ro.disconnect();
            make();
          }
        });
        ro.observe(holder);
      }
    }
    attempt(0);
  }

  // Robust CSV loader
  function loadCsvText(paths){
    var tried=[];
    function tryIdx(i){
      if(i>=paths.length){
        var err=new Error('CSV not found');
        err.tried=tried.slice();
        throw err;
      }
      var p=paths[i];
      return fetch(p,{cache:'no-store'}).then(function(r){
        tried.push(p+' ['+r.status+']');
        if(!r.ok)throw new Error('HTTP '+r.status);
        return r.text().then(function(txt){
          var source = p.indexOf('google') >= 0 ? 'Google Sheets' : 'Local file';
          console.log('✓ Loaded from ' + source + ': ' + p.substring(0, 80) + '...');
          return txt;
        });
      }).catch(function(){
        return tryIdx(i+1);
      });
    }
    return tryIdx(0);
  }

  /* ------------------------ Password Gate ------------------------ */
  var FINANCE_PASSWORD='Futuri123';
  function PasswordGate(props){
    var _ok=React.useState(function(){return sessionStorage.getItem('finance_ok')==='1';});
    var ok=_ok[0],setOk=_ok[1];
    var _pw=React.useState('');var pw=_pw[0],setPw=_pw[1];
    if(ok) return props.children;
    return e('div',{className:'max-w-sm mx-auto card p-5 space-y-3'},
      e('div',{className:'text-lg font-semibold'},'Enter password'),
      e('input',{type:'password',className:'w-full bg-[#0e141c] border border-slate-700 rounded-md px-3 py-2',placeholder:'Password',value:pw,onChange:function(ev){setPw(ev.target.value);}}),
      e('button',{className:'bg-slate-800 hover:bg-slate-700 rounded-md px-3 py-2 w-full',onClick:function(){if(pw===FINANCE_PASSWORD){sessionStorage.setItem('finance_ok','1');setOk(true);}else alert('Incorrect password');}},'Unlock'),
      e('div',{className:'text-xs text-slate-500'},'Protected area — client-side password only')
    );
  }

  /* ------------------------ Dashboard wrapper ------------------------ */
  var Dashboard = function Dashboard(props){
    return e('main', { className:'min-h-screen w-full bg-slate-950 text-slate-100 p-4' }, props.children);
  };

  /* ------------------------ KPI Screen ------------------------ */
  // Product order for display
  var ORDER = [
    'FUSION TOTAL','POST','STREAMING','SMARTSPEAKERS','MOBILE','LDR','AUDIOAI',
    'CONTENT INTELLIGENCE TOTAL','TOPICPULSE','PREP+',
    'SALES INTELLIGENCE TOTAL','SPOTON','TOPLINE',
    'TOTAL'
  ];
  
  var BOLD_KEYS = ['FUSION TOTAL','CONTENT INTELLIGENCE TOTAL','SALES INTELLIGENCE TOTAL','TOTAL','ALL PRODUCTS'];
  function isBoldKey(k){ return BOLD_KEYS.indexOf(k.toUpperCase())>=0; }

  function KPIScreen(){
    // State
    var _asOf = React.useState(function(){ return addMonths(new Date(),-1); });
    var asOf = _asOf[0], setAsOf = _asOf[1];
    var baseMonth = addMonths(asOf,-12);
    
    var _activeTab = React.useState('kpi-summary');
    var activeTab = _activeTab[0], setActiveTab = _activeTab[1];
    
    var _data = React.useState({ bob:[], cpmgi:[], activeBob:[], err:null, loading:true });
    var data = _data[0], setData = _data[1];

    // Active BoB filter state
    var _activeBobFilters = React.useState({
      parentCompany: '',
      market: '',
      station: '',
      product: '',
      ae: '',
      psmFusion: '',
      psmSi: '',
      psmCi: '',
      paymentMethod: '',
      contractRenewalStatus: ''
    });
    var activeBobFilters = _activeBobFilters[0], setActiveBobFilters = _activeBobFilters[1];

    // Active BoB sort state
    var _activeBobSort = React.useState({ column: 'parentCompany', direction: 'asc' });
    var activeBobSort = _activeBobSort[0], setActiveBobSort = _activeBobSort[1];
    
    // Load CSV data
    React.useEffect(function(){
      var bobPaths = ['/public/kpi_bob.csv','public/kpi_bob.csv','/kpi_bob.csv','kpi_bob.csv'];
      var cpmPaths = ['/public/kpi_cpmgi.csv','public/kpi_cpmgi.csv','/kpi_cpmgi.csv','kpi_cpmgi.csv'];
      var activeBobPaths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSx0Pp-5H60alDG7lTOneta10phn8QwLqXhnj0SSuAxobX5oaPj206IRFywm0BMBVPSOqCeKccE8KWY/pub?gid=414930090&single=true&output=csv',
        '/public/active_bob.csv','public/active_bob.csv','/active_bob.csv','active_bob.csv'
      ];

      Promise.all([
        loadCsvText(bobPaths).then(function(txt){ return { which:'bob', txt:txt }; }).catch(function(e){ return { which:'bob', err:e }; }),
        loadCsvText(cpmPaths).then(function(txt){ return { which:'cpmgi', txt:txt }; }).catch(function(e){ return { which:'cpmgi', err:e }; }),
        loadCsvText(activeBobPaths).then(function(txt){ return { which:'activeBob', txt:txt }; }).catch(function(e){ return { which:'activeBob', err:e }; })
      ]).then(function(results){
        var bobRes = results.find(function(r){ return r.which==='bob'; });
        var cpmRes = results.find(function(r){ return r.which==='cpmgi'; });
        var activeBobRes = results.find(function(r){ return r.which==='activeBob'; });

        var errors = [];
        if(bobRes.err) errors.push('kpi_bob.csv not found');
        if(cpmRes.err) errors.push('kpi_cpmgi.csv not found');
        // Active BoB is optional - don't error if not found

        if(errors.length){
          setData({ bob:[], cpmgi:[], activeBob:[], err: errors.join(', '), loading:false });
          return;
        }

        // Parse BoB data
        var bobParsed = Papa.parse(bobRes.txt, { header:true, skipEmptyLines:true });
        var bobRows = (bobParsed.data || []).map(function(r){
          var monthDate = parseMonth(r.Month);
          return {
            _m: monthDate,
            Month: r.Month,
            Product: (r.Product || '').toUpperCase().trim(),
            Product_Count: parseInt(r.Product_Count,10) || 0,
            Signed_Count: parseInt(r.Signed_Count,10) || 0,
            Activated_Count: parseInt(r.Activated_Count,10) || 0,
            Cancelled_Count: parseInt(r.Cancelled_Count,10) || 0,
            Signed_Value: parseMoney(r.Signed_Value),
            Activated_Value: parseMoney(r.Activated_Value),
            Cancelled_Value: parseMoney(r.Cancelled_Value),
            Product_Retention: r.Product_Retention === 'n/a' ? null : parseFloat(r.Product_Retention),
            Revenue: parseMoney(r.Revenue)
          };
        }).filter(function(r){ return r._m; });

        // Parse CPM data
        var cpmParsed = Papa.parse(cpmRes.txt, { header:true, skipEmptyLines:true });
        var cpmRows = (cpmParsed.data || []).filter(function(r){
          return r.Month && r.Month !== '2025'; // Filter out summary rows
        }).map(function(r){
          var monthDate = parseMonth(r.Month);
          var gi = String(r.Weekly_Gross_Impressions || '0').replace(/[,"]/g,'');
          return {
            _m: monthDate,
            Month: r.Month,
            CPM: parseFloat(r.CPM) || 0,
            Weekly_Gross_Impressions: parseFloat(gi) || 0
          };
        }).filter(function(r){ return r._m; });

        // Parse Active BoB data (if available)
        var activeBobRows = [];
        if(!activeBobRes.err && activeBobRes.txt){
          var activeBobParsed = Papa.parse(activeBobRes.txt, { header:true, skipEmptyLines:true });
          activeBobRows = (activeBobParsed.data || []).map(function(r){
            return {
              parentCompany: (r['Parent Company'] || r.parentCompany || '').trim(),
              market: (r['Market'] || r.market || '').trim(),
              station: (r['Station'] || r.station || '').trim(),
              product: (r['Product'] || r.product || '').trim(),
              ae: (r['AE'] || r.ae || '').trim(),
              psmFusion: (r['PSM Fusion'] || r.psmFusion || '').trim(),
              psmSi: (r['PSM SI'] || r.psmSi || '').trim(),
              psmCi: (r['PSM CI'] || r.psmCi || '').trim(),
              contractStartDate: (r['Contract Item Start Date'] || r.contractStartDate || '').trim(),
              contractEndDate: (r['Contract End Date'] || r.contractEndDate || '').trim(),
              paymentMethod: (r['Payment Method'] || r.paymentMethod || '').trim(),
              contractRenewalStatus: (r['Contract Renewal Status'] || r.contractRenewalStatus || '').trim(),
              annualContractValue: parseMoney(r['Annual Contract Value'] || r.annualContractValue || 0)
            };
          }).filter(function(r){ return r.parentCompany || r.station; });
        }

        // Find most recent month with data
        var latestMonth = null;
        bobRows.forEach(function(r){
          if(r.Product === 'TOTAL' || r.Product === 'ALL PRODUCTS'){
            if(!latestMonth || r._m > latestMonth) latestMonth = r._m;
          }
        });
        if(latestMonth) setAsOf(latestMonth);

        setData({ bob:bobRows, cpmgi:cpmRows, activeBob:activeBobRows, err:null, loading:false });
      }).catch(function(err){
        setData({ bob:[], cpmgi:[], activeBob:[], err: 'Error loading data: '+(err.message||err), loading:false });
      });
    }, []);
    
    var bob = data.bob;
    var cpmgi = data.cpmgi;
    var activeBob = data.activeBob;
    
    // Helper functions for data access
    function row(product, d){
      var p = (product || '').toUpperCase();
      if(p === 'TOTAL') p = 'TOTAL';
      return bob.find(function(r){
        var rp = r.Product;
        if(p === 'TOTAL' && (rp === 'TOTAL' || rp === 'ALL PRODUCTS')) return sameMonth(r._m, d);
        return rp === p && sameMonth(r._m, d);
      });
    }
    
    function productCountFor(product, d){
      var r = row(product, d);
      return r ? r.Product_Count : null;
    }
    
    function sumWhere(filterFn, valueFn){
      var total = 0;
      bob.forEach(function(r){
        if(filterFn(r)) total += (valueFn(r) || 0);
      });
      return total;
    }
    
    function retention(product, d){
      d = d || asOf;
      // Use Product_Retention from the CSV directly
      var retentionVal = null;
      bob.forEach(function(r){
        var p = r.Product;
        var targetP = (product || '').toUpperCase();
        if(targetP === 'TOTAL' && (p !== 'TOTAL' && p !== 'ALL PRODUCTS')) return;
        if(targetP !== 'TOTAL' && p !== targetP) return;
        if(!sameMonth(r._m, d)) return;
        if(r.Product_Retention != null && r.Product_Retention !== 'n/a'){
          retentionVal = r.Product_Retention;
        }
      });
      return retentionVal;
    }
    
    function cvr(product, d){
      d = d || asOf;
      var yr = d.getFullYear();
      
      // Get starting revenue (December of prior year)
      var priorYearEnd = new Date(yr - 1, 11, 31);
      var startingRevenue = 0;
      var cancelledValue = 0;
      
      // Find starting revenue from Dec of prior year
      bob.forEach(function(r){
        var p = r.Product;
        var targetP = (product || '').toUpperCase();
        if(targetP === 'TOTAL' && (p !== 'TOTAL' && p !== 'ALL PRODUCTS')) return;
        if(targetP !== 'TOTAL' && p !== targetP) return;
        
        // Get revenue from Dec of prior year as starting point
        if(r._m && r._m.getFullYear() === yr - 1 && r._m.getMonth() === 11){
          startingRevenue += (r.Revenue || 0);
        }
      });
      
      // Sum cancelled value YTD
      var startOfYear = new Date(yr, 0, 1);
      bob.forEach(function(r){
        var p = r.Product;
        var targetP = (product || '').toUpperCase();
        if(targetP === 'TOTAL' && (p !== 'TOTAL' && p !== 'ALL PRODUCTS')) return;
        if(targetP !== 'TOTAL' && p !== targetP) return;
        if(!r._m || r._m < startOfYear || r._m > d) return;
        cancelledValue += Math.abs(r.Cancelled_Value || 0);
      });
      
      if(startingRevenue === 0) return null;
      return (startingRevenue - cancelledValue) / startingRevenue;
    }
    
    function aggYTD(product, col, yr){
      var end = new Date(yr, asOf.getMonth(), 1);
      var start = new Date(yr, 0, 1);
      return sumWhere(function(r){
        var p = r.Product;
        var targetP = (product || '').toUpperCase();
        if(targetP === 'TOTAL' && (p !== 'TOTAL' && p !== 'ALL PRODUCTS')) return false;
        if(targetP !== 'TOTAL' && p !== targetP) return false;
        return r._m && r._m >= start && r._m <= end;
      }, function(r){
        if(col === 'Cancelled_Value') return Math.abs(r.Cancelled_Value || 0);
        return r[col] || 0;
      });
    }
    
    function aggByMonth(product, d, col){
      var total = 0;
      bob.forEach(function(r){
        var p = r.Product;
        var targetP = (product || '').toUpperCase();
        if(targetP === 'TOTAL' && (p !== 'TOTAL' && p !== 'ALL PRODUCTS')) return;
        if(targetP !== 'TOTAL' && p !== targetP) return;
        if(!sameMonth(r._m, d)) return;
        if(col === 'Cancelled_Value') total += Math.abs(r.Cancelled_Value || 0);
        else total += (r[col] || 0);
      });
      return total;
    }
    
    // Compute top metrics
    var topMetrics = ORDER.map(function(key){
      var current = productCountFor(key, asOf);
      var pyCount = productCountFor(key, baseMonth);
      var pmCount = productCountFor(key, addMonths(asOf,-1));
      var yoyRetention = retention(key, asOf);
      var cvrVal = cvr(key, asOf);
      var dMoMPct = (current != null && pmCount != null && pmCount !== 0) ? (current - pmCount) / pmCount : null;
      var dYoYPct = (current != null && pyCount != null && pyCount !== 0) ? (current - pyCount) / pyCount : null;
      return {
        key: key,
        name: key === 'TOTAL' ? 'All Products' : key,
        current: current,
        pyCount: pyCount,
        yoyRetention: yoyRetention,
        cvr: cvrVal,
        dMoMPct: dMoMPct,
        dYoYPct: dYoYPct
      };
    });
    
    // Build table data arrays
    function buildTableData(countCol, valueCol, goodUp){
      return ORDER.map(function(product){
        var mCY = aggByMonth(product, asOf, countCol);
        var mPY = aggByMonth(product, baseMonth, countCol);
        var yCY = aggYTD(product, countCol, asOf.getFullYear());
        var yPY = aggYTD(product, countCol, baseMonth.getFullYear());
        var mCh = mPY ? (mCY - mPY) / mPY : null;
        var yCh = yPY ? (yCY - yPY) / yPY : null;
        return { product: product, mCY:mCY, mPY:mPY, mCh:mCh, yCY:yCY, yPY:yPY, yCh:yCh, goodUp:goodUp };
      });
    }
    
    function buildMoneyTableData(col, goodUp){
      return ORDER.map(function(product){
        var mCY = aggByMonth(product, asOf, col);
        var mPY = aggByMonth(product, baseMonth, col);
        var yCY = aggYTD(product, col, asOf.getFullYear());
        var yPY = aggYTD(product, col, baseMonth.getFullYear());
        var mCh = mPY ? (mCY - mPY) / mPY : null;
        var yCh = yPY ? (yCY - yPY) / yPY : null;
        return { product: product, mCY:mCY, mPY:mPY, mCh:mCh, yCY:yCY, yPY:yPY, yCh:yCh, goodUp:goodUp };
      });
    }
    
    function buildNetNewData(){
      return ORDER.map(function(product){
        var actMCY = aggByMonth(product, asOf, 'Activated_Count');
        var canMCY = aggByMonth(product, asOf, 'Cancelled_Count');
        var actMPY = aggByMonth(product, baseMonth, 'Activated_Count');
        var canMPY = aggByMonth(product, baseMonth, 'Cancelled_Count');
        var actYCY = aggYTD(product, 'Activated_Count', asOf.getFullYear());
        var canYCY = aggYTD(product, 'Cancelled_Count', asOf.getFullYear());
        var actYPY = aggYTD(product, 'Activated_Count', baseMonth.getFullYear());
        var canYPY = aggYTD(product, 'Cancelled_Count', baseMonth.getFullYear());
        
        var mCY = actMCY - canMCY;
        var mPY = actMPY - canMPY;
        var yCY = actYCY - canYCY;
        var yPY = actYPY - canYPY;
        var mCh = mPY ? (mCY - mPY) / Math.abs(mPY) : null;
        var yCh = yPY ? (yCY - yPY) / Math.abs(yPY) : null;
        return { product: product, mCY:mCY, mPY:mPY, mCh:mCh, yCY:yCY, yPY:yPY, yCh:yCh, goodUp:true };
      });
    }
    
    function buildNetNew$Data(){
      return ORDER.map(function(product){
        var actMCY = aggByMonth(product, asOf, 'Activated_Value');
        var canMCY = aggByMonth(product, asOf, 'Cancelled_Value');
        var actMPY = aggByMonth(product, baseMonth, 'Activated_Value');
        var canMPY = aggByMonth(product, baseMonth, 'Cancelled_Value');
        var actYCY = aggYTD(product, 'Activated_Value', asOf.getFullYear());
        var canYCY = aggYTD(product, 'Cancelled_Value', asOf.getFullYear());
        var actYPY = aggYTD(product, 'Activated_Value', baseMonth.getFullYear());
        var canYPY = aggYTD(product, 'Cancelled_Value', baseMonth.getFullYear());
        
        var mCY = actMCY - canMCY;
        var mPY = actMPY - canMPY;
        var yCY = actYCY - canYCY;
        var yPY = actYPY - canYPY;
        var mCh = mPY ? (mCY - mPY) / Math.abs(mPY) : null;
        var yCh = yPY ? (yCY - yPY) / Math.abs(yPY) : null;
        return { product: product, mCY:mCY, mPY:mPY, mCh:mCh, yCY:yCY, yPY:yPY, yCh:yCh, goodUp:true };
      });
    }
    
    var signed = buildTableData('Signed_Count', null, true);
    var signed$ = buildMoneyTableData('Signed_Value', true);
    var acts = buildTableData('Activated_Count', null, true);
    var activated$ = buildMoneyTableData('Activated_Value', true);
    var canc = buildTableData('Cancelled_Count', null, false);
    var cancelled$ = buildMoneyTableData('Cancelled_Value', false);
    var netNew = buildNetNewData();
    var netNew$ = buildNetNew$Data();
    
    // Loading/Error states
    if(data.loading){
      return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 flex items-center justify-center'},
        e('div',{className:'text-slate-400'},'Loading KPI data...')
      );
    }
    if(data.err){
      return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6'},
        e('div',{className:'card p-6'},
          e('div',{className:'text-lg font-semibold text-rose-400 mb-2'},'Error Loading Data'),
          e('div',{className:'text-slate-400'},data.err)
        )
      );
    }

    /* ---------- KPI SUMMARY FUNCTIONS ---------- */
    function computeKPIMetrics() {
      var currentYear = asOf.getFullYear();
      var priorYear = currentYear - 1;
      
      var currentData = {}, priorData = {};
      
      ORDER.forEach(function(product) {
        var currentRow = row(product, asOf);
        var priorRow = row(product, new Date(priorYear, asOf.getMonth(), 1));
        
        currentData[product] = {
          count: productCountFor(product, asOf) || 0,
          revenue: currentRow ? parseMoney(currentRow.Revenue) || 0 : 0
        };
        
        priorData[product] = {
          count: productCountFor(product, new Date(priorYear, asOf.getMonth(), 1)) || 0,
          revenue: priorRow ? parseMoney(priorRow.Revenue) || 0 : 0
        };
      });
      
      return { current: currentData, prior: priorData };
    }

    function KPISummaryTables() {
      var kpiData = computeKPIMetrics();
      var current = kpiData.current;
      var prior = kpiData.prior;
      
      function ProductCountsTable() {
        function rowData(product) {
          var curr = current[product] ? current[product].count : 0;
          var prev = prior[product] ? prior[product].count : 0;
          var growth = prev ? (curr - prev) / prev : null;
          var currPct = current['TOTAL'] ? (curr / current['TOTAL'].count) * 100 : 0;
          var prevPct = prior['TOTAL'] ? (prev / prior['TOTAL'].count) * 100 : 0;
          var pctChange = currPct - prevPct;
          
          return { product: product, curr: curr, prev: prev, growth: growth, currPct: currPct, prevPct: prevPct, pctChange: pctChange };
        }
        
        var tableData = ORDER.map(rowData);
        
        return e('div', { className: 'grid grid-cols-1 lg:grid-cols-2 gap-4' }, [
          e('div', { key:'counts', className: 'card p-4' }, [
            e('h3', { className: 'text-sm font-semibold mb-3 text-slate-200' }, 'Product Counts'),
            e('div', { className: 'overflow-x-auto' }, 
              e('table', { className: 'w-full text-xs' }, [
                e('thead', {key:'th'},
                  e('tr', { className: 'text-[10px] text-slate-400 border-b border-slate-600' }, [
                    e('th', { key:'p', className: 'text-left py-2 px-1' }, 'Product'),
                    e('th', { key:'py', className: 'text-right py-2 px-1' }, monthShortLabel(addMonths(asOf, -12))),
                    e('th', { key:'cy', className: 'text-right py-2 px-1' }, monthShortLabel(asOf)),
                    e('th', { key:'g', className: 'text-right py-2 px-1' }, 'YoY Growth')
                  ])
                ),
                e('tbody', {key:'tb'},
                  tableData.map(function(r) {
                    var bold = isBoldKey(r.product);
                    return e('tr', { key: r.product, className: 'border-t border-slate-700/30 ' + (bold ? 'font-semibold' : '') }, [
                      e('td', { key:'p', className: 'py-1 px-1 text-left' }, r.product === 'TOTAL' ? 'Total' : r.product),
                      e('td', { key:'py', className: 'py-1 px-1 text-right' }, r.prev || 0),
                      e('td', { key:'cy', className: 'py-1 px-1 text-right' }, r.curr || 0),
                      e('td', { key:'g', className: 'py-1 px-1 text-right ' + clsDelta(r.growth, true) }, fmtPercent(r.growth))
                    ]);
                  })
                )
              ])
            )
          ]),
          
          e('div', { key:'mix', className: 'card p-4' }, [
            e('h3', { className: 'text-sm font-semibold mb-3 text-slate-200' }, 'Product Count Mix'),
            e('div', { className: 'overflow-x-auto' }, 
              e('table', { className: 'w-full text-xs' }, [
                e('thead', {key:'th'},
                  e('tr', { className: 'text-[10px] text-slate-400 border-b border-slate-600' }, [
                    e('th', { key:'p', className: 'text-left py-2 px-1' }, 'Product'),
                    e('th', { key:'py', className: 'text-right py-2 px-1' }, monthShortLabel(addMonths(asOf, -12))),
                    e('th', { key:'cy', className: 'text-right py-2 px-1' }, monthShortLabel(asOf)),
                    e('th', { key:'c', className: 'text-right py-2 px-1' }, 'Change in %')
                  ])
                ),
                e('tbody', {key:'tb'},
                  tableData.filter(function(r) { return r.product !== 'TOTAL'; }).map(function(r) {
                    var bold = isBoldKey(r.product);
                    return e('tr', { key: r.product, className: 'border-t border-slate-700/30 ' + (bold ? 'font-semibold' : '') }, [
                      e('td', { key:'p', className: 'py-1 px-1 text-left' }, r.product),
                      e('td', { key:'py', className: 'py-1 px-1 text-right' }, (r.prevPct || 0).toFixed(1) + '%'),
                      e('td', { key:'cy', className: 'py-1 px-1 text-right' }, (r.currPct || 0).toFixed(1) + '%'),
                      e('td', { key:'c', className: 'py-1 px-1 text-right ' + clsDelta(r.pctChange, true) }, 
                        (r.pctChange >= 0 ? '+' : '') + r.pctChange.toFixed(1) + '%')
                    ]);
                  })
                )
              ])
            )
          ])
        ]);
      }

      function ProductRevenueTable() {
        function rowData(product) {
          var curr = current[product] ? current[product].revenue : 0;
          var prev = prior[product] ? prior[product].revenue : 0;
          var growth = prev ? (curr - prev) / prev : null;
          var currPct = current['TOTAL'] ? (curr / current['TOTAL'].revenue) * 100 : 0;
          var prevPct = prior['TOTAL'] ? (prev / prior['TOTAL'].revenue) * 100 : 0;
          var pctChange = currPct - prevPct;
          
          return { product: product, curr: curr, prev: prev, growth: growth, currPct: currPct, prevPct: prevPct, pctChange: pctChange };
        }
        
        var tableData = ORDER.map(rowData);
        
        return e('div', { className: 'grid grid-cols-1 lg:grid-cols-2 gap-4' }, [
          e('div', { key:'rev', className: 'card p-4' }, [
            e('h3', { className: 'text-sm font-semibold mb-3 text-slate-200' }, 'Product Revenue (Annualized)'),
            e('div', { className: 'overflow-x-auto' }, 
              e('table', { className: 'w-full text-xs' }, [
                e('thead', {key:'th'},
                  e('tr', { className: 'text-[10px] text-slate-400 border-b border-slate-600' }, [
                    e('th', { key:'p', className: 'text-left py-2 px-1' }, 'Product'),
                    e('th', { key:'py', className: 'text-right py-2 px-1' }, monthShortLabel(addMonths(asOf, -12))),
                    e('th', { key:'cy', className: 'text-right py-2 px-1' }, monthShortLabel(asOf)),
                    e('th', { key:'g', className: 'text-right py-2 px-1' }, 'YoY Growth')
                  ])
                ),
                e('tbody', {key:'tb'},
                  tableData.map(function(r) {
                    var bold = isBoldKey(r.product);
                    return e('tr', { key: r.product, className: 'border-t border-slate-700/30 ' + (bold ? 'font-semibold' : '') }, [
                      e('td', { key:'p', className: 'py-1 px-1 text-left' }, r.product === 'TOTAL' ? 'Total' : r.product),
                      e('td', { key:'py', className: 'py-1 px-1 text-right' }, fmtCompactMoney(r.prev)),
                      e('td', { key:'cy', className: 'py-1 px-1 text-right' }, fmtCompactMoney(r.curr)),
                      e('td', { key:'g', className: 'py-1 px-1 text-right ' + clsDelta(r.growth, true) }, fmtPercent(r.growth))
                    ]);
                  })
                )
              ])
            )
          ]),
          
          e('div', { key:'mix', className: 'card p-4' }, [
            e('h3', { className: 'text-sm font-semibold mb-3 text-slate-200' }, 'Product Revenue Mix'),
            e('div', { className: 'overflow-x-auto' }, 
              e('table', { className: 'w-full text-xs' }, [
                e('thead', {key:'th'},
                  e('tr', { className: 'text-[10px] text-slate-400 border-b border-slate-600' }, [
                    e('th', { key:'p', className: 'text-left py-2 px-1' }, 'Product'),
                    e('th', { key:'py', className: 'text-right py-2 px-1' }, monthShortLabel(addMonths(asOf, -12))),
                    e('th', { key:'cy', className: 'text-right py-2 px-1' }, monthShortLabel(asOf)),
                    e('th', { key:'c', className: 'text-right py-2 px-1' }, 'Change in %')
                  ])
                ),
                e('tbody', {key:'tb'},
                  tableData.filter(function(r) { return r.product !== 'TOTAL'; }).map(function(r) {
                    var bold = isBoldKey(r.product);
                    return e('tr', { key: r.product, className: 'border-t border-slate-700/30 ' + (bold ? 'font-semibold' : '') }, [
                      e('td', { key:'p', className: 'py-1 px-1 text-left' }, r.product),
                      e('td', { key:'py', className: 'py-1 px-1 text-right' }, (r.prevPct || 0).toFixed(1) + '%'),
                      e('td', { key:'cy', className: 'py-1 px-1 text-right' }, (r.currPct || 0).toFixed(1) + '%'),
                      e('td', { key:'c', className: 'py-1 px-1 text-right ' + clsDelta(r.pctChange, true) }, 
                        (r.pctChange >= 0 ? '+' : '') + r.pctChange.toFixed(1) + '%')
                    ]);
                  })
                )
              ])
            )
          ])
        ]);
      }

      function AnnualRevenuePerProductTable() {
        function rowData(product) {
          var currRevenue = current[product] ? current[product].revenue : 0;
          var prevRevenue = prior[product] ? prior[product].revenue : 0;
          var currCount = current[product] ? current[product].count : 0;
          var prevCount = prior[product] ? prior[product].count : 0;
          
          var currRevPerProduct = currCount ? currRevenue / currCount : 0;
          var prevRevPerProduct = prevCount ? prevRevenue / prevCount : 0;
          var growth = prevRevPerProduct ? (currRevPerProduct - prevRevPerProduct) / prevRevPerProduct : null;
          
          return { product: product, curr: currRevPerProduct, prev: prevRevPerProduct, growth: growth };
        }
        
        var tableData = ORDER.map(rowData);
        
        return e('div', { className: 'card p-4' }, [
          e('h3', { className: 'text-sm font-semibold mb-3 text-slate-200' }, 'Annual Revenue Per Product'),
          e('div', { className: 'overflow-x-auto' }, 
            e('table', { className: 'w-full text-xs' }, [
              e('thead', {key:'th'},
                e('tr', { className: 'text-[10px] text-slate-400 border-b border-slate-600' }, [
                  e('th', { key:'p', className: 'text-left py-2 px-1' }, 'Product'),
                  e('th', { key:'py', className: 'text-right py-2 px-1' }, monthShortLabel(addMonths(asOf, -12))),
                  e('th', { key:'cy', className: 'text-right py-2 px-1' }, monthShortLabel(asOf)),
                  e('th', { key:'g', className: 'text-right py-2 px-1' }, 'YoY Growth')
                ])
              ),
              e('tbody', {key:'tb'},
                tableData.map(function(r) {
                  var bold = isBoldKey(r.product);
                  return e('tr', { key: r.product, className: 'border-t border-slate-700/30 ' + (bold ? 'font-semibold' : '') }, [
                    e('td', { key:'p', className: 'py-1 px-1 text-left' }, r.product === 'TOTAL' ? 'Weighted Total' : r.product),
                    e('td', { key:'py', className: 'py-1 px-1 text-right' }, fmtCompactMoney(r.prev)),
                    e('td', { key:'cy', className: 'py-1 px-1 text-right' }, fmtCompactMoney(r.curr)),
                    e('td', { key:'g', className: 'py-1 px-1 text-right ' + clsDelta(r.growth, true) }, fmtPercent(r.growth))
                  ]);
                })
              )
            ])
          )
        ]);
      }

      function RevenuePerProductPerMonthTable() {
        function rowData(product) {
          var currRevenue = current[product] ? current[product].revenue : 0;
          var prevRevenue = prior[product] ? prior[product].revenue : 0;
          var currCount = current[product] ? current[product].count : 0;
          var prevCount = prior[product] ? prior[product].count : 0;
          
          var currRevPerProductMonth = currCount ? (currRevenue / currCount) / 12 : 0;
          var prevRevPerProductMonth = prevCount ? (prevRevenue / prevCount) / 12 : 0;
          var growth = prevRevPerProductMonth ? (currRevPerProductMonth - prevRevPerProductMonth) / prevRevPerProductMonth : null;
          
          return { product: product, curr: currRevPerProductMonth, prev: prevRevPerProductMonth, growth: growth };
        }
        
        var tableData = ORDER.map(rowData);
        
        return e('div', { className: 'card p-4' }, [
          e('h3', { className: 'text-sm font-semibold mb-3 text-slate-200' }, 'Revenue Per Product Per Month'),
          e('div', { className: 'overflow-x-auto' }, 
            e('table', { className: 'w-full text-xs' }, [
              e('thead', {key:'th'},
                e('tr', { className: 'text-[10px] text-slate-400 border-b border-slate-600' }, [
                  e('th', { key:'p', className: 'text-left py-2 px-1' }, 'Product'),
                  e('th', { key:'py', className: 'text-right py-2 px-1' }, monthShortLabel(addMonths(asOf, -12))),
                  e('th', { key:'cy', className: 'text-right py-2 px-1' }, monthShortLabel(asOf)),
                  e('th', { key:'g', className: 'text-right py-2 px-1' }, 'YoY Growth')
                ])
              ),
              e('tbody', {key:'tb'},
                tableData.map(function(r) {
                  var bold = isBoldKey(r.product);
                  return e('tr', { key: r.product, className: 'border-t border-slate-700/30 ' + (bold ? 'font-semibold' : '') }, [
                    e('td', { key:'p', className: 'py-1 px-1 text-left' }, r.product === 'TOTAL' ? 'Weighted Total' : r.product),
                    e('td', { key:'py', className: 'py-1 px-1 text-right' }, fmtCompactMoney(r.prev)),
                    e('td', { key:'cy', className: 'py-1 px-1 text-right' }, fmtCompactMoney(r.curr)),
                    e('td', { key:'g', className: 'py-1 px-1 text-right ' + clsDelta(r.growth, true) }, fmtPercent(r.growth))
                  ]);
                })
              )
            ])
          )
        ]);
      }

      return e('section', { className: 'space-y-6' }, [
        e('h2', { key:'h', className: 'text-lg lg:text-xl font-bold text-slate-200 border-b border-slate-600 pb-2' }, 'Product Counts & Revenue Summary'),
        e(ProductCountsTable, {key:'pc'}),
        e(ProductRevenueTable, {key:'pr'}), 
        e(AnnualRevenuePerProductTable, {key:'ar'}),
        e(RevenuePerProductPerMonthTable, {key:'rm'})
      ]);
    }

    /* ---------- Table Block Components ---------- */
    function TableBlock(title, rows){
      function cell(n,goodUp){ 
        return e('td',{className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap '+clsDelta(n,goodUp)}, 
          prettyDelta(n)); 
      }
      function header(){ 
        return e('thead',{className:'sticky top-0 bg-slate-800 z-20'},
          e('tr',{className:'text-[10px] lg:text-[11px] text-slate-400 border-b border-slate-600'},
            e('th',{key:'p',className:'px-2 py-2 text-left sticky left-0 bg-slate-800 z-30 min-w-[100px] lg:min-w-[140px] border-r border-slate-600'},'Product'),
            e('th',{key:'mpy',className:'px-1 py-2 text-right min-w-[50px] lg:min-w-[65px]'},monthShortLabel(baseMonth)),
            e('th',{key:'mcy',className:'px-1 py-2 text-right min-w-[50px] lg:min-w-[65px]'},monthShortLabel(asOf)),
            e('th',{key:'md',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},'% Δ'),
            e('th',{key:'ypy',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px] border-l border-slate-600'},(baseMonth.getFullYear())+' YTD'),
            e('th',{key:'ycy',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},(asOf.getFullYear())+' YTD'),
            e('th',{key:'yd',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},'% Δ')
          )
        );
      }
      function body(){
        return e('tbody',null, rows.map(function(r){
          var bold = isBoldKey(r.product) || r.product.toUpperCase()==='TOTAL';
          var rowCls='border-t border-slate-700/30 hover:bg-slate-700/20 '+(bold?'font-semibold bg-slate-700/10':'');
          return e('tr',{key:r.product, className:rowCls},
            e('td',{key:'p',className:'px-2 py-2 text-left sticky left-0 bg-inherit z-10 text-xs lg:text-sm border-r border-slate-700/30'}, 
              r.product==='TOTAL'?'All Products':r.product),
            e('td',{key:'mpy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap'}, r.mPY||0),
            e('td',{key:'mcy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap'}, r.mCY||0),
            cell(r.mCh, r.goodUp),
            e('td',{key:'ypy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap border-l border-slate-700/30'}, r.yPY||0),
            e('td',{key:'ycy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap'}, r.yCY||0),
            cell(r.yCh, r.goodUp)
          );
        }));
      }
      return e('section',{className:'card p-3 lg:p-4 space-y-3'},
        e('div',{className:'flex items-center justify-between'},
          e('div',{className:'text-sm lg:text-base font-semibold'},title),
          e('div',{className:'text-xs text-slate-400'},monthShortLabel(asOf))
        ),
        e('div',{className:'overflow-x-auto rounded-lg border border-slate-700/50 -mx-1'},
          e('table',{className:'w-full text-xs lg:text-sm border-collapse'}, header(), body())
        )
      );
    }

    function MoneyTableBlock(title, rows){
      function cell(n,goodUp){ 
        return e('td',{className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap '+clsDelta(n,goodUp)}, 
          prettyDelta(n)); 
      }
      function header(){ 
        return e('thead',{className:'sticky top-0 bg-slate-800 z-20'},
          e('tr',{className:'text-[10px] lg:text-[11px] text-slate-400 border-b border-slate-600'},
            e('th',{key:'p',className:'px-2 py-2 text-left sticky left-0 bg-slate-800 z-30 min-w-[100px] lg:min-w-[140px] border-r border-slate-600'},'Product'),
            e('th',{key:'mpy',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},monthShortLabel(baseMonth)),
            e('th',{key:'mcy',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},monthShortLabel(asOf)),
            e('th',{key:'md',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},'% Δ'),
            e('th',{key:'ypy',className:'px-1 py-2 text-right min-w-[70px] lg:min-w-[85px] border-l border-slate-600'},(baseMonth.getFullYear())+' YTD'),
            e('th',{key:'ycy',className:'px-1 py-2 text-right min-w-[70px] lg:min-w-[85px]'},(asOf.getFullYear())+' YTD'),
            e('th',{key:'yd',className:'px-1 py-2 text-right min-w-[60px] lg:min-w-[75px]'},'% Δ')
          )
        );
      }
      function body(){
        return e('tbody',null, rows.map(function(r){
          var bold = isBoldKey(r.product) || r.product.toUpperCase()==='TOTAL';
          var rowCls='border-t border-slate-700/30 hover:bg-slate-700/20 '+(bold?'font-semibold bg-slate-700/10':'');
          return e('tr',{key:r.product, className:rowCls},
            e('td',{key:'p',className:'px-2 py-2 text-left sticky left-0 bg-inherit z-10 text-xs lg:text-sm border-r border-slate-700/30'}, 
              r.product==='TOTAL'?'All Products':r.product),
            e('td',{key:'mpy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap'}, fmtCompactMoney(r.mPY)),
            e('td',{key:'mcy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap'}, fmtCompactMoney(r.mCY)),
            cell(r.mCh, r.goodUp),
            e('td',{key:'ypy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap border-l border-slate-700/30'}, fmtCompactMoney(r.yPY)),
            e('td',{key:'ycy',className:'px-1 py-2 text-right text-xs lg:text-sm whitespace-nowrap'}, fmtCompactMoney(r.yCY)),
            cell(r.yCh, r.goodUp)
          );
        }));
      }
      return e('section',{className:'card p-3 lg:p-4 space-y-3'},
        e('div',{className:'flex items-center justify-between'},
          e('div',{className:'text-sm lg:text-base font-semibold'},title),
          e('div',{className:'text-xs text-slate-400'},monthShortLabel(asOf))
        ),
        e('div',{className:'overflow-x-auto rounded-lg border border-slate-700/50 -mx-1'},
          e('table',{className:'w-full text-xs lg:text-sm border-collapse'}, header(), body())
        )
      );
    }

    function ControlsKPI(){
      return e('section',{className:'card p-4 flex flex-wrap items-center gap-3 justify-between'},
        [
          e('div',{key:'left',className:'flex items-center gap-2'},[
            e('div',{key:'t',className:'text-base font-semibold'},'KPI Report — '+monthLongLabel(asOf)),
            e('span',{key:'p',className:'pill'},'Reporting Month: '+monthLongLabel(asOf))
          ]),
          e('div',{key:'right',className:'flex items-center gap-2'},[
            e('button',{key:'prev',className:'bg-slate-800 hover:bg-slate-700 rounded-md px-2 py-1 text-sm',
                        onClick:function(){ setAsOf(addMonths(asOf,-1)); }},'←'),
            e('input',{key:'inp',type:'month',className:'bg-[#0e141c] border border-slate-700 rounded-md px-2 py-1 text-sm',
                       value:monthKey(asOf),
                       onChange:function(ev){ setAsOf(firstOfMonth(ev.target.value)); }}),
            e('button',{key:'next',className:'bg-slate-800 hover:bg-slate-700 rounded-md px-2 py-1 text-sm',
                        onClick:function(){ setAsOf(addMonths(asOf,1)); }},'→')
          ])
        ]
      );
    }

    function TabNavigation() {
      var tabs = [
        { id: 'kpi-summary', label: 'Product Counts & Revenue', icon: '📊' },
        { id: 'kpi-analytics', label: 'KPI Analytics', icon: '📈' },
        { id: 'book-of-business', label: 'Book of Business', icon: '📋' },
        { id: 'cpm-gi', label: 'CPM & GI', icon: '📺' }
      ];

      return e('section', { className: 'card p-2' }, 
        e('div', { className: 'flex space-x-1 overflow-x-auto' }, 
          tabs.map(function(tab) {
            var isActive = activeTab === tab.id;
            var buttonClass = 'flex items-center gap-2 px-3 py-2 rounded-md text-sm font-medium transition-colors whitespace-nowrap ' +
              (isActive 
                ? 'bg-blue-600 text-white' 
                : 'text-slate-300 hover:text-white hover:bg-slate-700');
            
            return e('button', {
              key: tab.id,
              className: buttonClass,
              onClick: function() { setActiveTab(tab.id); }
            }, [
              e('span', {key:'i'}, tab.icon),
              e('span', {key:'l'}, tab.label)
            ]);
          })
        )
      );
    }

    function TopSummary(){
      function Th(txt){return e('th',{className:'px-2 py-1 text-[11px] text-slate-400 text-right'},txt);}
      return e('section',{className:'card p-4 space-y-2'},
        e('div',{className:'flex items-center justify-between'},
          e('div',{className:'text-sm font-semibold'},'Book of Business — '+monthLongLabel(asOf)),
          e('div',{className:'text-xs text-slate-400'},
            'Reporting Month: '+monthLongLabel(asOf)
          )
        ),
        e('div',{className:'overflow-auto'},
          e('table',{className:'min-w-[880px] text-sm'},
            e('thead',null,
              e('tr',{className:'text-[11px] text-slate-400'},
                e('th',{key:'p',className:'px-2 py-1 text-left'},'Product'),
                Th('Current Product Count'),
                Th('Product Retention (YoY)'),
                Th('Contract Value Retention (YTD)'),
                Th('Δ vs Prior Month (%)'),
                Th('Δ vs Prior Year (%)'),
                Th('PY Product Count')
              )
            ),
            e('tbody',null,
              topMetrics.map(function(m){
                var rowCls='border-t border-slate-700/40 '+(isBoldKey(m.key)?'font-semibold':'');
                return e('tr',{key:m.name,className:rowCls},
                  e('td',{key:'p',className:'px-2 py-1 text-left'},m.name),
                  e('td',{key:'c',className:'px-2 py-1 text-right'}, m.current==null?'—':m.current),
                  e('td',{key:'r',className:'px-2 py-1 text-right'}, fmtPercent(m.yoyRetention)),
                  e('td',{key:'v',className:'px-2 py-1 text-right'}, fmtPercent(m.cvr)),
                  e('td',{key:'mom',className:'px-2 py-1 text-right '+clsDelta(m.dMoMPct,true)}, fmtPercent(m.dMoMPct)),
                  e('td',{key:'yoy',className:'px-2 py-1 text-right '+clsDelta(m.dYoYPct,true)}, fmtPercent(m.dYoYPct)),
                  e('td',{key:'py',className:'px-2 py-1 text-right'}, m.pyCount==null?'—':m.pyCount)
                );
              })
            )
          )
        )
      );
    }

    function BookOfBusinessTabContent() {
      return e('div', { className: 'space-y-8' }, [
        e('section', { key:'bob', className: 'space-y-4' }, [
          e('h2', { key:'h', className: 'text-lg lg:text-xl font-bold text-slate-200 border-b border-slate-600 pb-2' }, 'Book of Business Overview'),
          e(TopSummary, {key:'ts'})
        ]),

        e('section',{key:'signed',className:'space-y-4'},
          e('h2',{className:'text-lg lg:text-xl font-bold text-slate-200 border-b border-slate-600 pb-2'},'Products Signed'),
          e('div',{className:'grid grid-cols-1 2xl:grid-cols-2 gap-6'},[
            TableBlock('Count', signed),
            MoneyTableBlock('Value ($)', signed$)
          ])
        ),

        e('section',{key:'activated',className:'space-y-4'},
          e('h2',{className:'text-lg lg:text-xl font-bold text-slate-200 border-b border-slate-600 pb-2'},'Products Activated'),
          e('div',{className:'grid grid-cols-1 2xl:grid-cols-2 gap-6'},[
            TableBlock('Count', acts),
            MoneyTableBlock('Value ($)', activated$)
          ])
        ),

        e('section',{key:'cancelled',className:'space-y-4'},
          e('h2',{className:'text-lg lg:text-xl font-bold text-slate-200 border-b border-slate-600 pb-2'},'Products Cancelled'),
          e('div',{className:'grid grid-cols-1 2xl:grid-cols-2 gap-6'},[
            TableBlock('Count', canc),
            MoneyTableBlock('Value ($)', cancelled$)
          ])
        ),

        e('section',{key:'netnew',className:'space-y-4'},
          e('h2',{className:'text-lg lg:text-xl font-bold text-slate-200 border-b border-slate-600 pb-2'},'Net New Business'),
          e('div',{className:'grid grid-cols-1 2xl:grid-cols-2 gap-6'},[
            TableBlock('Count (Activated − Cancelled)', netNew),
            MoneyTableBlock('Value (Activated − Cancelled)', netNew$)
          ])
        )
      ]);
    }

    function KPIAnalyticsTabContent() {
      var currentMonth = asOf;
      var priorMonth = addMonths(asOf, -1);
      var priorYear = addMonths(asOf, -12);
      
      var currentCount = productCountFor('TOTAL', currentMonth) || 0;
      var priorMonthCount = productCountFor('TOTAL', priorMonth) || 0;
      var priorYearCount = productCountFor('TOTAL', priorYear) || 0;
      
      var currentRetention = retention('TOTAL');
      var priorMonthRetention = retention('TOTAL', priorMonth);
      var priorYearRetention = retention('TOTAL', priorYear);
      
      var currentCVR = cvr('TOTAL');
      var priorYearCVR = cvr('TOTAL', priorYear);
      
      var currentYearYTD = asOf.getFullYear();
      var priorYearYTD = currentYearYTD - 1;
      
      var signedYTD = aggYTD('TOTAL', 'Signed_Count', currentYearYTD);
      var signedYTDPY = aggYTD('TOTAL', 'Signed_Count', priorYearYTD);
      var activatedYTD = aggYTD('TOTAL', 'Activated_Count', currentYearYTD);
      var activatedYTDPY = aggYTD('TOTAL', 'Activated_Count', priorYearYTD);
      var cancelledYTD = aggYTD('TOTAL', 'Cancelled_Count', currentYearYTD);
      var cancelledYTDPY = aggYTD('TOTAL', 'Cancelled_Count', priorYearYTD);
      var netNewYTD = activatedYTD - cancelledYTD;
      var netNewYTDPY = activatedYTDPY - cancelledYTDPY;
      
      var signedValueYTD = aggYTD('TOTAL', 'Signed_Value', currentYearYTD);
      var signedValueYTDPY = aggYTD('TOTAL', 'Signed_Value', priorYearYTD);
      var activatedValueYTD = aggYTD('TOTAL', 'Activated_Value', currentYearYTD);
      var activatedValueYTDPY = aggYTD('TOTAL', 'Activated_Value', priorYearYTD);
      var cancelledValueYTD = aggYTD('TOTAL', 'Cancelled_Value', currentYearYTD);
      var cancelledValueYTDPY = aggYTD('TOTAL', 'Cancelled_Value', priorYearYTD);
      var netNewValueYTD = activatedValueYTD - cancelledValueYTD;
      var netNewValueYTDPY = activatedValueYTDPY - cancelledValueYTDPY;
      
      var totalRow = row('TOTAL', currentMonth);
      var priorMonthRow = row('TOTAL', priorMonth);
      var priorYearRow = row('TOTAL', priorYear);
      
      var currentRevenue = totalRow ? parseMoney(totalRow.Revenue) || 0 : 0;
      var priorMonthRevenue = priorMonthRow ? parseMoney(priorMonthRow.Revenue) || 0 : 0;
      var priorYearRevenue = priorYearRow ? parseMoney(priorYearRow.Revenue) || 0 : 0;
      
      var revenuePerProduct = currentCount ? currentRevenue / currentCount : 0;
      var revenuePerProductPM = priorMonthCount ? priorMonthRevenue / priorMonthCount : 0;
      var revenuePerProductPY = priorYearCount ? priorYearRevenue / priorYearCount : 0;
      
      function AnalyticsTile(props) {
        var pmVariance = (props.priorMonth !== null && props.priorMonth !== undefined) ? deltaPct(props.current, props.priorMonth) : null;
        var pyVariance = (props.priorYear !== null && props.priorYear !== undefined) ? deltaPct(props.current, props.priorYear) : null;
        
        function formatValue(val) {
          if (props.isPct) return fmtPercent(val);
          if (props.isMoney) return fmtCompactMoney(val);
          return Math.round(val || 0).toString();
        }
        
        return e('div', { className: 'card p-4 hover:shadow-lg transition-shadow' }, [
          e('div', { key:'t', className: 'text-xs font-medium text-slate-400 uppercase tracking-wider mb-2' }, props.title),
          e('div', { key:'v', className: 'text-2xl font-bold text-slate-100 mb-3' }, formatValue(props.current)),
          e('div', { key:'c', className: 'space-y-1' }, 
            [
              (pmVariance !== null) ? e('div', { key:'pm', className: 'flex justify-between items-center text-sm' }, [
                e('span', { key:'l', className: 'text-slate-500' }, 'vs Prior Month:'),
                e('span', { key:'v', className: 'font-medium ' + clsDelta(pmVariance, props.goodUp || true) }, (pmVariance >= 0 ? '+' : '') + fmtPercent(pmVariance))
              ]) : null,
              (pyVariance !== null) ? e('div', { key:'py', className: 'flex justify-between items-center text-sm' }, [
                e('span', { key:'l', className: 'text-slate-500' }, 'vs Prior Year:'),
                e('span', { key:'v', className: 'font-medium ' + clsDelta(pyVariance, props.goodUp || true) }, (pyVariance >= 0 ? '+' : '') + fmtPercent(pyVariance))
              ]) : null
            ].filter(Boolean)
          )
        ]);
      }
      
      return e('div', { className: 'space-y-8' }, [
        e('section', { key:'current', className: 'space-y-4' }, [
          e('h2', { className: 'text-xl font-bold text-slate-200 border-b border-slate-600 pb-2' }, 'Key Performance Indicators'),
          e('h3', { className: 'text-lg font-semibold text-slate-300 mb-3' }, 'Current Month Metrics'),
          e('div', { className: 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4' }, [
            e(AnalyticsTile, { key:'cnt', title: 'Total Product Count', current: currentCount, priorMonth: priorMonthCount, priorYear: priorYearCount, goodUp: true }),
            e(AnalyticsTile, { key:'ret', title: 'Product Retention (YoY)', current: currentRetention, priorMonth: priorMonthRetention, priorYear: priorYearRetention, isPct: true, goodUp: true }),
            e(AnalyticsTile, { key:'cvr', title: 'Contract Value Retention', current: currentCVR, priorMonth: null, priorYear: priorYearCVR, isPct: true, goodUp: true }),
            e(AnalyticsTile, { key:'rpp', title: 'Revenue Per Product (Annual)', current: revenuePerProduct, priorMonth: revenuePerProductPM, priorYear: revenuePerProductPY, isMoney: true, goodUp: true })
          ])
        ]),
        
        e('section', { key:'ytdcount', className: 'space-y-4' }, [
          e('h3', { className: 'text-lg font-semibold text-slate-300 mb-3' }, 'Year-to-Date Count Metrics'),
          e('div', { className: 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4' }, [
            e(AnalyticsTile, { key:'sig', title: 'Products Signed YTD', current: signedYTD, priorMonth: null, priorYear: signedYTDPY, goodUp: true }),
            e(AnalyticsTile, { key:'act', title: 'Products Activated YTD', current: activatedYTD, priorMonth: null, priorYear: activatedYTDPY, goodUp: true }),
            e(AnalyticsTile, { key:'can', title: 'Products Cancelled YTD', current: cancelledYTD, priorMonth: null, priorYear: cancelledYTDPY, goodUp: false }),
            e(AnalyticsTile, { key:'net', title: 'Net New Products YTD', current: netNewYTD, priorMonth: null, priorYear: netNewYTDPY, goodUp: true })
          ])
        ]),

        e('section', { key:'ytdvalue', className: 'space-y-4' }, [
          e('h3', { className: 'text-lg font-semibold text-slate-300 mb-3' }, 'Year-to-Date Value Metrics'),
          e('div', { className: 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4' }, [
            e(AnalyticsTile, { key:'sigv', title: 'Value Signed YTD', current: signedValueYTD, priorMonth: null, priorYear: signedValueYTDPY, isMoney: true, goodUp: true }),
            e(AnalyticsTile, { key:'actv', title: 'Value Activated YTD', current: activatedValueYTD, priorMonth: null, priorYear: activatedValueYTDPY, isMoney: true, goodUp: true }),
            e(AnalyticsTile, { key:'canv', title: 'Value Cancelled YTD', current: cancelledValueYTD, priorMonth: null, priorYear: cancelledValueYTDPY, isMoney: true, goodUp: false }),
            e(AnalyticsTile, { key:'netv', title: 'Net New Value YTD', current: netNewValueYTD, priorMonth: null, priorYear: netNewValueYTDPY, isMoney: true, goodUp: true })
          ])
        ])
      ]);
    }

    function CPMGITabContent() {
      if (!cpmgi || !cpmgi.length) {
        return e('div', { className: 'card p-8 text-center' }, [
          e('h3', { key:'h', className: 'text-lg font-semibold mb-3 text-slate-200' }, 'CPM & Gross Impressions'),
          e('div', { key:'m', className: 'text-slate-400' }, 'No CPM data found. Check if kpi_cpmgi.csv is loaded.')
        ]);
      }
      
      var currentCPM = cpmgi.find(function(r) { return sameMonth(r._m, asOf); });
      var priorCPM = cpmgi.find(function(r) { return sameMonth(r._m, addMonths(asOf, -12)); });
      var priorMonthCPM = cpmgi.find(function(r) { return sameMonth(r._m, addMonths(asOf, -1)); });
      
      var currCPMVal = currentCPM ? currentCPM.CPM : 0;
      var prevCPMVal = priorCPM ? priorCPM.CPM : 0;
      var prevMonthCPMVal = priorMonthCPM ? priorMonthCPM.CPM : 0;
      var cpmYoYGrowth = prevCPMVal ? (currCPMVal - prevCPMVal) / prevCPMVal : null;
      var cpmMoMGrowth = prevMonthCPMVal ? (currCPMVal - prevMonthCPMVal) / prevMonthCPMVal : null;
      
      var currGI = currentCPM ? (currentCPM.Weekly_Gross_Impressions || 0) * 4.33 : 0;
      var prevGI = priorCPM ? (priorCPM.Weekly_Gross_Impressions || 0) * 4.33 : 0;
      var prevMonthGI = priorMonthCPM ? (priorMonthCPM.Weekly_Gross_Impressions || 0) * 4.33 : 0;
      var giYoYGrowth = prevGI ? (currGI - prevGI) / prevGI : null;
      var giMoMGrowth = prevMonthGI ? (currGI - prevMonthGI) / prevMonthGI : null;
      
      return e('div', { className: 'space-y-6' }, [
        e('div', { key:'tables', className: 'grid grid-cols-1 lg:grid-cols-2 gap-6' }, [
          e('div', { key:'cpm', className: 'card p-6' }, [
            e('h3', { className: 'text-lg font-semibold mb-4 text-slate-200 border-b border-slate-600 pb-2' }, 'CPM Performance'),
            e('div', { className: 'overflow-x-auto' }, 
              e('table', { className: 'w-full text-sm' }, [
                e('thead', {key:'th'},
                  e('tr', { className: 'text-xs text-slate-400 border-b border-slate-600' }, [
                    e('th', { key:'m', className: 'text-left py-3 px-2' }, 'Metric'),
                    e('th', { key:'pm', className: 'text-right py-3 px-2' }, monthShortLabel(addMonths(asOf, -1))),
                    e('th', { key:'py', className: 'text-right py-3 px-2' }, monthShortLabel(addMonths(asOf, -12))),
                    e('th', { key:'cy', className: 'text-right py-3 px-2' }, monthShortLabel(asOf)),
                    e('th', { key:'mom', className: 'text-right py-3 px-2' }, 'MoM %'),
                    e('th', { key:'yoy', className: 'text-right py-3 px-2' }, 'YoY %')
                  ])
                ),
                e('tbody', {key:'tb'}, [
                  e('tr', { key:'cpm', className: 'border-t border-slate-700/30' }, [
                    e('td', { key:'m', className: 'py-3 px-2 text-left font-medium' }, 'CPM'),
                    e('td', { key:'pm', className: 'py-3 px-2 text-right' }, '$' + (prevMonthCPMVal || 0).toFixed(2)),
                    e('td', { key:'py', className: 'py-3 px-2 text-right' }, '$' + (prevCPMVal || 0).toFixed(2)),
                    e('td', { key:'cy', className: 'py-3 px-2 text-right font-semibold' }, '$' + (currCPMVal || 0).toFixed(2)),
                    e('td', { key:'mom', className: 'py-3 px-2 text-right font-medium ' + clsDelta(cpmMoMGrowth, true) }, fmtPercent(cpmMoMGrowth)),
                    e('td', { key:'yoy', className: 'py-3 px-2 text-right font-medium ' + clsDelta(cpmYoYGrowth, true) }, fmtPercent(cpmYoYGrowth))
                  ])
                ])
              ])
            )
          ]),
          
          e('div', { key:'gi', className: 'card p-6' }, [
            e('h3', { className: 'text-lg font-semibold mb-4 text-slate-200 border-b border-slate-600 pb-2' }, 'Gross Impressions (Monthly)'),
            e('div', { className: 'overflow-x-auto' }, 
              e('table', { className: 'w-full text-sm' }, [
                e('thead', {key:'th'},
                  e('tr', { className: 'text-xs text-slate-400 border-b border-slate-600' }, [
                    e('th', { key:'m', className: 'text-left py-3 px-2' }, 'Metric'),
                    e('th', { key:'pm', className: 'text-right py-3 px-2' }, monthShortLabel(addMonths(asOf, -1))),
                    e('th', { key:'py', className: 'text-right py-3 px-2' }, monthShortLabel(addMonths(asOf, -12))),
                    e('th', { key:'cy', className: 'text-right py-3 px-2' }, monthShortLabel(asOf)),
                    e('th', { key:'mom', className: 'text-right py-3 px-2' }, 'MoM %'),
                    e('th', { key:'yoy', className: 'text-right py-3 px-2' }, 'YoY %')
                  ])
                ),
                e('tbody', {key:'tb'}, [
                  e('tr', { key:'gi', className: 'border-t border-slate-700/30' }, [
                    e('td', { key:'m', className: 'py-3 px-2 text-left font-medium' }, 'Gross Impressions'),
                    e('td', { key:'pm', className: 'py-3 px-2 text-right' }, fmtGIThousands(prevMonthGI)),
                    e('td', { key:'py', className: 'py-3 px-2 text-right' }, fmtGIThousands(prevGI)),
                    e('td', { key:'cy', className: 'py-3 px-2 text-right font-semibold' }, fmtGIThousands(currGI)),
                    e('td', { key:'mom', className: 'py-3 px-2 text-right font-medium ' + clsDelta(giMoMGrowth, true) }, fmtPercent(giMoMGrowth)),
                    e('td', { key:'yoy', className: 'py-3 px-2 text-right font-medium ' + clsDelta(giYoYGrowth, true) }, fmtPercent(giYoYGrowth))
                  ])
                ])
              ])
            )
          ])
        ]),
        
        e('div', { key:'notes', className: 'card p-6' }, [
          e('h3', { className: 'text-lg font-semibold mb-3 text-slate-200' }, 'Notes'),
          e('div', { className: 'text-sm text-slate-400 space-y-2' }, [
            e('div', {key:'1'}, '• Weekly Gross Impressions converted to monthly using 4.33x multiplier (52 weeks ÷ 12 months)'),
            e('div', {key:'2'}, '• CPM values show cost per thousand impressions (displayed to 2 decimal places)'),
            e('div', {key:'3'}, '• MoM = Month-over-Month, YoY = Year-over-Year growth rates')
          ])
        ])
      ]);
    }

    // Active Book of Business Tab Content
    function ActiveBookOfBusinessTabContent() {
      // Get unique values for filters
      function getUniqueValues(field) {
        var vals = {};
        activeBob.forEach(function(r){ if(r[field]) vals[r[field]] = true; });
        return Object.keys(vals).sort();
      }

      // Filter dropdown component
      function FilterSelect(props) {
        return e('select', {
          className: 'bg-slate-700 border border-slate-600 rounded px-2 py-1 text-xs text-slate-200 min-w-[120px]',
          value: activeBobFilters[props.field] || '',
          onChange: function(ev) {
            var newFilters = Object.assign({}, activeBobFilters);
            newFilters[props.field] = ev.target.value;
            setActiveBobFilters(newFilters);
          }
        }, [
          e('option', { key: '', value: '' }, 'All ' + props.label),
          getUniqueValues(props.field).map(function(v) {
            return e('option', { key: v, value: v }, v);
          })
        ]);
      }

      // Apply filters
      var filteredData = activeBob.filter(function(r) {
        if (activeBobFilters.parentCompany && r.parentCompany !== activeBobFilters.parentCompany) return false;
        if (activeBobFilters.market && r.market !== activeBobFilters.market) return false;
        if (activeBobFilters.station && r.station !== activeBobFilters.station) return false;
        if (activeBobFilters.product && r.product !== activeBobFilters.product) return false;
        if (activeBobFilters.ae && r.ae !== activeBobFilters.ae) return false;
        if (activeBobFilters.psmFusion && r.psmFusion !== activeBobFilters.psmFusion) return false;
        if (activeBobFilters.psmSi && r.psmSi !== activeBobFilters.psmSi) return false;
        if (activeBobFilters.psmCi && r.psmCi !== activeBobFilters.psmCi) return false;
        if (activeBobFilters.paymentMethod && r.paymentMethod !== activeBobFilters.paymentMethod) return false;
        if (activeBobFilters.contractRenewalStatus && r.contractRenewalStatus !== activeBobFilters.contractRenewalStatus) return false;
        return true;
      });

      // Apply sorting
      var sortedData = filteredData.slice().sort(function(a, b) {
        var aVal = a[activeBobSort.column] || '';
        var bVal = b[activeBobSort.column] || '';
        if (activeBobSort.column === 'annualContractValue') {
          aVal = a.annualContractValue || 0;
          bVal = b.annualContractValue || 0;
        }
        if (aVal < bVal) return activeBobSort.direction === 'asc' ? -1 : 1;
        if (aVal > bVal) return activeBobSort.direction === 'asc' ? 1 : -1;
        return 0;
      });

      // Calculate KPIs
      var totalACV = filteredData.reduce(function(sum, r) { return sum + (r.annualContractValue || 0); }, 0);
      var activeContracts = filteredData.length;
      var uniqueParents = {};
      var uniqueStations = {};
      filteredData.forEach(function(r) {
        if (r.parentCompany) uniqueParents[r.parentCompany] = true;
        if (r.station) uniqueStations[r.station] = true;
      });

      // Sort handler
      function handleSort(column) {
        if (activeBobSort.column === column) {
          setActiveBobSort({ column: column, direction: activeBobSort.direction === 'asc' ? 'desc' : 'asc' });
        } else {
          setActiveBobSort({ column: column, direction: 'asc' });
        }
      }

      // Sort indicator
      function SortIndicator(props) {
        if (activeBobSort.column !== props.column) return null;
        return e('span', { className: 'ml-1' }, activeBobSort.direction === 'asc' ? '▲' : '▼');
      }

      // Sortable header
      function SortableHeader(props) {
        return e('th', {
          key: props.column,
          className: 'px-2 py-2 text-left text-xs font-medium text-slate-400 cursor-pointer hover:text-slate-200 whitespace-nowrap',
          onClick: function() { handleSort(props.column); }
        }, [props.label, e(SortIndicator, { key: 's', column: props.column })]);
      }

      // Clear filters
      function clearFilters() {
        setActiveBobFilters({
          parentCompany: '', market: '', station: '', product: '', ae: '',
          psmFusion: '', psmSi: '', psmCi: '', paymentMethod: '', contractRenewalStatus: ''
        });
      }

      if (!activeBob.length) {
        return e('div', { className: 'card p-6 text-center text-slate-400' },
          'No Active Book of Business data available. Upload active_bob.csv to the public folder.'
        );
      }

      return e('div', { className: 'space-y-4' }, [
        // KPI Cards
        e('div', { key: 'kpis', className: 'grid grid-cols-2 lg:grid-cols-4 gap-4' }, [
          e('div', { key: 'acv', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Total ACV'),
            e('div', { className: 'text-2xl font-bold text-emerald-400' }, fmtCompact(totalACV))
          ]),
          e('div', { key: 'contracts', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Active Contracts'),
            e('div', { className: 'text-2xl font-bold text-blue-400' }, activeContracts.toLocaleString())
          ]),
          e('div', { key: 'parents', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Parent Companies'),
            e('div', { className: 'text-2xl font-bold text-purple-400' }, Object.keys(uniqueParents).length.toLocaleString())
          ]),
          e('div', { key: 'stations', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Stations'),
            e('div', { className: 'text-2xl font-bold text-orange-400' }, Object.keys(uniqueStations).length.toLocaleString())
          ])
        ]),

        // Filters
        e('div', { key: 'filters', className: 'card p-4' }, [
          e('div', { className: 'flex items-center justify-between mb-3' }, [
            e('h3', { key: 'h', className: 'text-sm font-semibold text-slate-300' }, 'Filters'),
            e('button', {
              key: 'clear',
              className: 'text-xs text-blue-400 hover:text-blue-300',
              onClick: clearFilters
            }, 'Clear All')
          ]),
          e('div', { className: 'flex flex-wrap gap-2' }, [
            e(FilterSelect, { key: 'f1', field: 'parentCompany', label: 'Parent Company' }),
            e(FilterSelect, { key: 'f2', field: 'market', label: 'Market' }),
            e(FilterSelect, { key: 'f3', field: 'station', label: 'Station' }),
            e(FilterSelect, { key: 'f4', field: 'product', label: 'Product' }),
            e(FilterSelect, { key: 'f5', field: 'ae', label: 'AE' }),
            e(FilterSelect, { key: 'f6', field: 'psmFusion', label: 'PSM Fusion' }),
            e(FilterSelect, { key: 'f7', field: 'psmSi', label: 'PSM SI' }),
            e(FilterSelect, { key: 'f8', field: 'psmCi', label: 'PSM CI' }),
            e(FilterSelect, { key: 'f9', field: 'paymentMethod', label: 'Payment Method' }),
            e(FilterSelect, { key: 'f10', field: 'contractRenewalStatus', label: 'Renewal Status' })
          ])
        ]),

        // Data Table
        e('div', { key: 'table', className: 'card p-4' }, [
          e('div', { className: 'flex items-center justify-between mb-3' }, [
            e('h3', { key: 'h', className: 'text-sm font-semibold text-slate-300' }, 'Active Contracts'),
            e('span', { key: 'count', className: 'text-xs text-slate-400' }, sortedData.length + ' records')
          ]),
          e('div', { className: 'overflow-x-auto' },
            e('table', { className: 'w-full text-sm' }, [
              e('thead', { key: 'thead' },
                e('tr', { className: 'border-b border-slate-600' }, [
                  e(SortableHeader, { key: 'h1', column: 'parentCompany', label: 'Parent Company' }),
                  e(SortableHeader, { key: 'h2', column: 'market', label: 'Market' }),
                  e(SortableHeader, { key: 'h3', column: 'station', label: 'Station' }),
                  e(SortableHeader, { key: 'h4', column: 'product', label: 'Product' }),
                  e(SortableHeader, { key: 'h5', column: 'ae', label: 'AE' }),
                  e(SortableHeader, { key: 'h6', column: 'psmFusion', label: 'PSM Fusion' }),
                  e(SortableHeader, { key: 'h7', column: 'psmSi', label: 'PSM SI' }),
                  e(SortableHeader, { key: 'h8', column: 'psmCi', label: 'PSM CI' }),
                  e(SortableHeader, { key: 'h9', column: 'contractStartDate', label: 'Start Date' }),
                  e(SortableHeader, { key: 'h10', column: 'contractEndDate', label: 'End Date' }),
                  e(SortableHeader, { key: 'h11', column: 'paymentMethod', label: 'Payment' }),
                  e(SortableHeader, { key: 'h12', column: 'contractRenewalStatus', label: 'Renewal Status' }),
                  e(SortableHeader, { key: 'h13', column: 'annualContractValue', label: 'ACV' })
                ])
              ),
              e('tbody', { key: 'tbody' },
                sortedData.slice(0, 500).map(function(r, i) {
                  return e('tr', { key: i, className: 'border-t border-slate-700/30 hover:bg-slate-700/20' }, [
                    e('td', { key: 'c1', className: 'px-2 py-2 whitespace-nowrap' }, r.parentCompany),
                    e('td', { key: 'c2', className: 'px-2 py-2 whitespace-nowrap' }, r.market),
                    e('td', { key: 'c3', className: 'px-2 py-2 whitespace-nowrap' }, r.station),
                    e('td', { key: 'c4', className: 'px-2 py-2 whitespace-nowrap' }, r.product),
                    e('td', { key: 'c5', className: 'px-2 py-2 whitespace-nowrap' }, r.ae),
                    e('td', { key: 'c6', className: 'px-2 py-2 whitespace-nowrap' }, r.psmFusion),
                    e('td', { key: 'c7', className: 'px-2 py-2 whitespace-nowrap' }, r.psmSi),
                    e('td', { key: 'c8', className: 'px-2 py-2 whitespace-nowrap' }, r.psmCi),
                    e('td', { key: 'c9', className: 'px-2 py-2 whitespace-nowrap' }, r.contractStartDate),
                    e('td', { key: 'c10', className: 'px-2 py-2 whitespace-nowrap' }, r.contractEndDate),
                    e('td', { key: 'c11', className: 'px-2 py-2 whitespace-nowrap' }, r.paymentMethod),
                    e('td', { key: 'c12', className: 'px-2 py-2 whitespace-nowrap' }, r.contractRenewalStatus),
                    e('td', { key: 'c13', className: 'px-2 py-2 whitespace-nowrap text-right font-medium text-emerald-400' }, fmtCurrency(r.annualContractValue))
                  ]);
                })
              )
            ])
          ),
          sortedData.length > 500 ? e('div', { className: 'text-xs text-slate-400 mt-2 text-center' }, 'Showing first 500 of ' + sortedData.length + ' records') : null
        ])
      ]);
    }

    function renderTabContent() {
      switch (activeTab) {
        case 'kpi-summary':
          return e(KPISummaryTables);
        case 'kpi-analytics':
          return e(KPIAnalyticsTabContent);
        case 'book-of-business':
          return e(BookOfBusinessTabContent);
        case 'cpm-gi':
          return e(CPMGITabContent);
        default:
          return e(KPISummaryTables);
      }
    }

    return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800'},
      e('div',{className:'max-w-[2000px] mx-auto px-4 lg:px-6 xl:px-8 py-6 space-y-6'},
        e(ControlsKPI),
        e(TabNavigation),
        e('div', { className: 'min-h-[400px]' }, renderTabContent()),
        e('footer',{className:'text-xs text-slate-500 text-center py-4 border-t border-slate-700'},
          'Data refreshed: '+monthLongLabel(asOf)+' • Navigate between tabs to explore different KPI views'
        )
      )
    );
  }

  /* ------------------------ Financials Screen ------------------------ */
  var pad2 = function(n){ return (n < 10 ? '0' : '') + n; };

  var monthKeyFromAnything = function(v){
    if(v == null) return null;
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
      return v.getFullYear() + '-' + pad2(v.getMonth() + 1);
    }
    var s = String(v).trim();
    if(!s) return null;
    var m;
    m = s.match(/^(\d{1,2})\s*\/\s*(\d{4})$/);
    if(m){
      var mm1 = parseInt(m[1],10);
      var yy1 = parseInt(m[2],10);
      if(mm1>=1 && mm1<=12) return yy1 + '-' + pad2(mm1);
    }
    m = s.match(/^(\d{4})\s*-\s*(\d{1,2})/);
    if(m){
      var yy2 = parseInt(m[1],10);
      var mm2 = parseInt(m[2],10);
      if(mm2>=1 && mm2<=12) return yy2 + '-' + pad2(mm2);
    }
    m = s.match(/^(\d{4})\s*\/\s*(\d{1,2})/);
    if(m){
      var yy3 = parseInt(m[1],10);
      var mm3 = parseInt(m[2],10);
      if(mm3>=1 && mm3<=12) return yy3 + '-' + pad2(mm3);
    }
    var t = Date.parse(s);
    if(!isNaN(t)){
      var d = new Date(t);
      return d.getFullYear() + '-' + pad2(d.getMonth()+1);
    }
    return null;
  };

  var toNumber = function(x){
    if(x == null) return null;
    if(typeof x === 'number') return isFinite(x) ? x : null;
    var s = String(x).trim();
    if(!s) return null;
    s = s.replace(/[$,]/g,'');
    var n = parseFloat(s);
    return isFinite(n) ? n : null;
  };

  function FinancialsScreen(){
    var _state = React.useState({ data:null, err:null, loading:true, loadTime:null, dataSource:null });
    var state = _state[0], setState = _state[1];
    
    var _viewMode = React.useState('monthly');
    var viewMode = _viewMode[0], setViewMode = _viewMode[1];
    
    var _selectedMonth = React.useState(null);
    var selectedMonth = _selectedMonth[0], setSelectedMonth = _selectedMonth[1];
    
    var _showCOGSDetail = React.useState(false);
    var showCOGSDetail = _showCOGSDetail[0], setShowCOGSDetail = _showCOGSDetail[1];
    
    var _showOPEXDetail = React.useState(false);
    var showOPEXDetail = _showOPEXDetail[0], setShowOPEXDetail = _showOPEXDetail[1];

    var REVENUE_ITEMS = ['Premiere', 'Other Barter', 'Cash', 'Misc.'];
    var COGS_ITEMS = ['Network', 'COGS Personnel', 'Dev and Support Tools', 'Webhosting', 'Other COGS'];
    var OPEX_ITEMS = ['Personnel', 'Benefits', 'Bonus', 'Recruiting', 'Office, Insurance, and Misc Exp', 'Professional Fees', 'Travel', 'Company Meetings', 'Marketing and Trade Shows', 'Systems'];

    React.useEffect(function(){
      // Cache-buster to force fresh data on each load
      var cacheBuster = '&_t=' + Date.now();
      
      // Google Sheets published URLs (primary), local files (fallback)
      var aPaths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSICXr4bYp-4kszCotiAIHXGBJElALUawJZBRit5fvs2fRD9DQZZSTN9lusoAkdazXDHKvjjh2n1sbB/pub?gid=1420819606&single=true&output=csv' + cacheBuster,
        '/public/actual.csv','public/actual.csv','/actual.csv','actual.csv'
      ];
      var bPaths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSICXr4bYp-4kszCotiAIHXGBJElALUawJZBRit5fvs2fRD9DQZZSTN9lusoAkdazXDHKvjjh2n1sbB/pub?gid=586410538&single=true&output=csv' + cacheBuster,
        '/public/budget.csv','public/budget.csv','/budget.csv','budget.csv'
      ];
      var fPaths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSICXr4bYp-4kszCotiAIHXGBJElALUawJZBRit5fvs2fRD9DQZZSTN9lusoAkdazXDHKvjjh2n1sbB/pub?gid=1345288878&single=true&output=csv' + cacheBuster,
        '/public/forecast.csv','public/forecast.csv','/forecast.csv','forecast.csv'
      ];

      Promise.all([
        loadCsvText(aPaths).then(function(txt){return {which:'actual',txt:txt};}).catch(function(er){return {which:'actual',err:er};}),
        loadCsvText(bPaths).then(function(txt){return {which:'budget',txt:txt};}).catch(function(er){return {which:'budget',err:er};}),
        loadCsvText(fPaths).then(function(txt){return {which:'forecast',txt:txt};}).catch(function(er){return {which:'forecast',err:er};})
      ]).then(function(parts){
        var pa = parts.find(function(p){return p.which==='actual';});
        var pb = parts.find(function(p){return p.which==='budget';});
        var pf = parts.find(function(p){return p.which==='forecast';});

        if(pa.err && pb.err && pf.err){
          setState({ data:null, err: 'Could not load any financial CSV files', loading:false });
          return;
        }

        function parseRows(txt){
          if(!txt) return [];
          // Clean up Windows line endings
          txt = txt.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
          var rows = Papa.parse(txt, {header:true, skipEmptyLines:true}).data || [];
          if(rows.length === 0) return [];
          
          var firstRow = rows[0];
          var keys = Object.keys(firstRow);
          
          // Detect format: row-based has "Line Item" as first column
          var isRowBased = keys[0] === 'Line Item';
          
          if(isRowBased){
            // Row-based format: Line Item, 1/2026, 2/2026, ...
            // Need to transpose to column-based format
            console.log('Detected row-based format, transposing...');
            
            // Get month columns (skip Line Item)
            var monthCols = keys.filter(function(k){ return /^\d{1,2}\/\d{4}$/.test(k); });
            
            // Create one row per month
            return monthCols.map(function(monthKey){
              var normalized = { monthKey: monthKey };
              rows.forEach(function(row){
                var lineItem = (row['Line Item'] || '').trim();
                if(lineItem){
                  var val = parseFloat(String(row[monthKey] || '0').replace(/[$,]/g,''));
                  normalized[lineItem] = isNaN(val) ? 0 : val;
                }
              });
              return normalized;
            });
          } else {
            // Column-based format: first column is empty (monthKey), line items as columns
            return rows.map(function(row){
              var monthKey = (row[''] || '').trim();
              var normalized = { monthKey: monthKey };
              Object.keys(row).forEach(function(k){
                if(k !== ''){
                  var cleanKey = k.trim();
                  var rawVal = row[k];
                  var val = parseFloat(String(rawVal).replace(/[$,\r\n]/g,'').trim());
                  normalized[cleanKey] = isNaN(val) ? 0 : val;
                }
              });
              return normalized;
            });
          }
        }

        var actualData = parseRows(pa.txt);
        var budgetData = parseRows(pb.txt);
        var forecastData = parseRows(pf.txt);
        
        // Debug: log structure
        if(actualData.length > 0){
          console.log('Actual data loaded:', actualData.length, 'rows');
          console.log('Sample keys:', Object.keys(actualData[0]));
        }
        
        var lastActualMonth = null;
        var actualMonths = [];
        actualData.forEach(function(r){
          var totalRev = r['Total Rev'] || 0;
          if(totalRev > 100){
            lastActualMonth = r.monthKey;
            actualMonths.push(r.monthKey);
          }
        });
        
        console.log('Actual months found:', actualMonths.length, 'months -', actualMonths.join(', '));
        console.log('Last actual month:', lastActualMonth);

        setState({ 
          data: { actual: actualData, budget: budgetData, forecast: forecastData, lastActualMonth: lastActualMonth, actualMonths: actualMonths }, 
          err:null, loading:false,
          loadTime: new Date().toLocaleString(),
          dataSource: 'Google Sheets'
        });
        
        if(lastActualMonth) setSelectedMonth(lastActualMonth);
      }).catch(function(er){
        setState({ data:null, err:'Error: '+(er&&er.message?er.message:String(er)), loading:false });
      });
    },[]);

    function parseMonthKey(mk){
      if(!mk) return null;
      var parts = mk.split('/');
      if(parts.length !== 2) return null;
      return { month: parseInt(parts[0],10), year: parseInt(parts[1],10) };
    }
    function getMonthNum(mk){ var p = parseMonthKey(mk); return p ? p.month : 0; }
    function getYear(mk){ var p = parseMonthKey(mk); return p ? p.year : 0; }
    function formatMonthKey(mk){
      var p = parseMonthKey(mk);
      if(!p) return mk;
      var names = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return names[p.month] + ' ' + p.year;
    }
    function getRow(dataset, monthKey){ return dataset.find(function(r){ return r.monthKey === monthKey; }) || null; }
    function getVal(dataset, monthKey, field){ var row = getRow(dataset, monthKey); return row ? (row[field] || 0) : 0; }
    
    function calcLTM(dataset, throughMonthKey, field, actualMonths){
      var p = parseMonthKey(throughMonthKey);
      if(!p) {
        console.log('calcLTM: invalid monthKey', throughMonthKey);
        return 0;
      }
      
      // Generate 12 month keys ending at throughMonthKey
      var ltmKeys = [];
      for(var i = 0; i < 12; i++){
        var m = p.month - i;
        var y = p.year;
        while(m <= 0){ m += 12; y -= 1; }
        ltmKeys.push(m + '/' + y);
      }
      
      var sum = 0;
      var matchCount = 0;
      
      // Iterate through dataset and sum values for matching months
      for(var j = 0; j < dataset.length; j++){
        var row = dataset[j];
        var rowKey = row.monthKey;
        
        // Check if this month is in our LTM range
        var inLTM = false;
        for(var k = 0; k < ltmKeys.length; k++){
          if(ltmKeys[k] === rowKey){
            inLTM = true;
            break;
          }
        }
        
        // Check if this is a real data row (Total Rev > 100)
        var isRealData = (row['Total Rev'] || 0) > 100;
        
        if(inLTM && isRealData){
          var val = row[field];
          if(typeof val === 'number' && !isNaN(val)){
            sum += val;
            matchCount++;
          }
        }
      }
      
      console.log('calcLTM(' + field + ') through ' + throughMonthKey + ': matched ' + matchCount + ' months, sum=' + sum.toFixed(2));
      return sum;
    }
    
    // Calculate Prior Year LTM (12 months ending same month last year)
    function calcPYLTM(dataset, throughMonthKey, field, actualMonths){
      var p = parseMonthKey(throughMonthKey);
      if(!p) return 0;
      var pyMonthKey = p.month + '/' + (p.year - 1);
      return calcLTM(dataset, pyMonthKey, field, actualMonths);
    }
    
    // Calculate Prior Year YTD
    function calcPYYTD(dataset, throughMonthKey, field, actualMonths){
      var p = parseMonthKey(throughMonthKey);
      if(!p) return 0;
      var targetYear = p.year - 1;
      var targetMonth = p.month;
      var sum = 0;
      dataset.forEach(function(row){
        if(getYear(row.monthKey) === targetYear && getMonthNum(row.monthKey) <= targetMonth && actualMonths.indexOf(row.monthKey) >= 0){
          sum += (row[field] || 0);
        }
      });
      return sum;
    }
    
    // Get Prior Year monthly value
    function getPYMonthly(dataset, monthKey, field, actualMonths){
      var p = parseMonthKey(monthKey);
      if(!p) return 0;
      var pyKey = p.month + '/' + (p.year - 1);
      if(actualMonths.indexOf(pyKey) >= 0){
        return getVal(dataset, pyKey, field);
      }
      return 0;
    }
    
    function calcYTD(dataset, throughMonthKey, field){
      var targetYear = getYear(throughMonthKey);
      var targetMonth = getMonthNum(throughMonthKey);
      var sum = 0;
      dataset.forEach(function(row){
        if(getYear(row.monthKey) === targetYear && getMonthNum(row.monthKey) <= targetMonth){
          sum += (row[field] || 0);
        }
      });
      return sum;
    }
    
    function calcFY(dataset, year, field){
      var sum = 0;
      dataset.forEach(function(row){ if(getYear(row.monthKey) === year){ sum += (row[field] || 0); } });
      return sum;
    }

    if(state.loading){
      return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6'},
        e('div',{className:'card p-6'}, e('div',{className:'text-lg font-semibold'},'Financials'), e('div',{className:'text-slate-400'},'Loading…'))
      );
    }
    if(state.err){
      return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6'},
        e('div',{className:'card p-6'}, e('div',{className:'text-lg font-semibold mb-2 text-red-400'},'Error'), e('pre',{className:'text-xs text-rose-300'},String(state.err)))
      );
    }

    var fin = state.data;
    if(!selectedMonth) return null;
    var selYear = getYear(selectedMonth);
    var selMonthNum = getMonthNum(selectedMonth);
    
    // Safety check
    if(!selYear || selYear < 2000){
      console.error('Invalid selYear:', selYear, 'from selectedMonth:', selectedMonth);
      selYear = 2025;
    }
    
    function getValue(dataset, field){
      if(viewMode === 'monthly') return getVal(dataset, selectedMonth, field);
      if(viewMode === 'ytd') return calcYTD(dataset, selectedMonth, field);
      return calcLTM(dataset, selectedMonth, field, fin.actualMonths);
    }
    function getBudgetValue(field){
      // Budget only makes sense for YTD/Monthly in 2026
      if(viewMode === 'ltm') return null; // No budget for LTM
      if(selYear !== 2026) return null; // No budget for 2025
      if(viewMode === 'monthly'){ return getVal(fin.budget, selMonthNum + '/2026', field); }
      if(viewMode === 'ytd'){
        var sum = 0;
        fin.budget.forEach(function(row){ if(getMonthNum(row.monthKey) <= selMonthNum){ sum += (row[field] || 0); } });
        return sum;
      }
      return calcFY(fin.budget, 2026, field);
    }
    function getForecastValue(field){
      // Forecast only makes sense for YTD/Monthly in 2026
      if(viewMode === 'ltm') return null;
      if(selYear !== 2026) return null;
      if(viewMode === 'monthly'){ return getVal(fin.forecast, selMonthNum + '/2026', field); }
      if(viewMode === 'ytd'){
        var sum = 0;
        fin.forecast.forEach(function(row){ if(getMonthNum(row.monthKey) <= selMonthNum){ sum += (row[field] || 0); } });
        return sum;
      }
      return calcFY(fin.forecast, 2026, field);
    }
    function getPYValue(field){
      if(viewMode === 'monthly') return getPYMonthly(fin.actual, selectedMonth, field, fin.actualMonths);
      if(viewMode === 'ytd') return calcPYYTD(fin.actual, selectedMonth, field, fin.actualMonths);
      return calcPYLTM(fin.actual, selectedMonth, field, fin.actualMonths);
    }
    
    // Check if we have PY data
    var hasPYData = fin.actualMonths.some(function(mk){ return getYear(mk) === selYear - 1; });
    // Check if we have budget (only for 2026)
    var hasBudget = selYear === 2026 && viewMode !== 'ltm';

    function SummaryCards(){
      var actualRev = getValue(fin.actual, 'Total Rev');
      var actualGM = getValue(fin.actual, 'Gross Margin');
      var actualEBITDA = getValue(fin.actual, 'EBITDA');
      
      // For LTM or when no budget, compare to PY
      var compareTobudget = hasBudget;
      var compRev, compGM, compEBITDA, compLabel;
      
      if(compareTobudget){
        compRev = getBudgetValue('Total Rev');
        compGM = getBudgetValue('Gross Margin');
        compEBITDA = getBudgetValue('EBITDA');
        compLabel = 'vs Budget';
      } else {
        compRev = getPYValue('Total Rev');
        compGM = getPYValue('Gross Margin');
        compEBITDA = getPYValue('EBITDA');
        compLabel = 'vs PY';
      }
      
      var revVar = compRev ? actualRev - compRev : null;
      var gmVar = compGM ? actualGM - compGM : null;
      var ebitdaVar = compEBITDA ? actualEBITDA - compEBITDA : null;
      var actualGMPct = actualRev ? actualGM / actualRev : 0;
      var actualEBITDAPct = actualRev ? actualEBITDA / actualRev : 0;

      function Card(props){
        var hasComparison = props.variance !== null && props.comparison;
        var varColor = props.variance >= 0 ? 'text-emerald-400' : 'text-rose-400';
        var statusColor = !hasComparison ? 'bg-slate-500' : (props.variance >= 0 ? 'bg-emerald-500' : 'bg-rose-500');
        return e('div',{className:'card p-5'},
          e('div',{className:'flex items-center justify-between mb-3'},
            e('div',{className:'text-xs font-medium text-slate-400 uppercase tracking-wider'}, props.title),
            e('div',{className:'w-2 h-2 rounded-full ' + statusColor})
          ),
          e('div',{className:'text-2xl font-bold text-slate-100 mb-1'}, props.value),
          props.pct ? e('div',{className:'text-sm text-slate-400 mb-2'}, props.pct) : null,
          hasComparison ? e('div',{className:'text-xs ' + varColor},
            (props.variance >= 0 ? '+' : '') + fmtCompact(props.variance) + ' ' + compLabel + ' (' + fmtPercent(props.comparison ? props.variance / props.comparison : 0) + ')'
          ) : e('div',{className:'text-xs text-slate-500'}, 'No comparison data')
        );
      }
      
      return e('div',{className:'grid grid-cols-1 md:grid-cols-3 gap-4'},
        e(Card,{title:'Total Revenue', value: fmtCompact(actualRev), variance: revVar, comparison: compRev}),
        e(Card,{title:'Gross Margin', value: fmtCompact(actualGM), pct: fmtPercent(actualGMPct) + ' GM', variance: gmVar, comparison: compGM}),
        e(Card,{title:'EBITDA', value: fmtCompact(actualEBITDA), pct: fmtPercent(actualEBITDAPct) + ' margin', variance: ebitdaVar, comparison: compEBITDA})
      );
    }

    function PLTable(){
      var showBudgetCols = hasBudget;
      var compLabel = showBudgetCols ? 'vs Budget' : 'vs PY';
      
      function PLRow(label, field, indent, isBold, isCost, isPercent){
        var actual = getValue(fin.actual, field);
        var budget = showBudgetCols ? getBudgetValue(field) : null;
        var forecast = showBudgetCols ? getForecastValue(field) : null;
        var py = hasPYData ? getPYValue(field) : null;
        
        var compValue = showBudgetCols ? budget : py;
        var variance = compValue !== null ? actual - compValue : null;
        var varColor = '';
        if(variance !== null){
          if(isCost){
            varColor = variance <= 0 ? 'text-emerald-400' : 'text-rose-400';
          } else {
            varColor = variance >= 0 ? 'text-emerald-400' : 'text-rose-400';
          }
        }
        var rowBg = isBold ? 'bg-slate-700/30' : '';
        var fontWeight = isBold ? 'font-semibold' : '';
        var paddingLeft = indent ? 'pl-6' : 'pl-2';
        
        var cells = [
          e('td',{key:'label',className:'py-2 ' + paddingLeft + ' ' + fontWeight}, label)
        ];
        
        if(showBudgetCols){
          cells.push(e('td',{key:'budget',className:'py-2 px-3 text-right ' + fontWeight}, isPercent ? fmtPercent(budget) : fmtCompact(budget || 0)));
          cells.push(e('td',{key:'forecast',className:'py-2 px-3 text-right ' + fontWeight}, isPercent ? fmtPercent(forecast) : fmtCompact(forecast || 0)));
        }
        
        cells.push(e('td',{key:'actual',className:'py-2 px-3 text-right ' + fontWeight}, isPercent ? fmtPercent(actual) : fmtCompact(actual)));
        
        if(hasPYData && selYear === 2026){
          cells.push(e('td',{key:'py',className:'py-2 px-3 text-right text-slate-400 ' + fontWeight}, isPercent ? fmtPercent(py) : fmtCompact(py || 0)));
        }
        
        cells.push(e('td',{key:'var',className:'py-2 px-3 text-right ' + varColor + ' ' + fontWeight}, 
          isPercent || variance === null ? '—' : (variance >= 0 ? '+' : '') + fmtCompact(variance)));
        cells.push(e('td',{key:'varpct',className:'py-2 px-3 text-right ' + varColor + ' ' + fontWeight}, 
          isPercent || variance === null || !compValue ? '—' : fmtPercent(variance / compValue)));
        
        return e('tr',{key:label, className:'border-t border-slate-700/30 ' + rowBg}, cells);
      }
      
      var tableRows = [];
      tableRows.push(PLRow('Total Revenue', 'Total Rev', false, true, false, false));
      REVENUE_ITEMS.forEach(function(item){ tableRows.push(PLRow(item, item, true, false, false, false)); });
      
      // COGS expandable header
      var cogsActual = getValue(fin.actual, 'Total COGS');
      var cogsBudget = showBudgetCols ? getBudgetValue('Total COGS') : null;
      var cogsForecast = showBudgetCols ? getForecastValue('Total COGS') : null;
      var cogsPY = hasPYData ? getPYValue('Total COGS') : null;
      var cogsComp = showBudgetCols ? cogsBudget : cogsPY;
      var cogsVar = cogsComp !== null ? cogsActual - cogsComp : null;
      var cogsVarColor = cogsVar !== null ? (cogsVar <= 0 ? 'text-emerald-400' : 'text-rose-400') : '';
      
      var cogsCells = [
        e('td',{key:'label',className:'py-2 pl-2 font-semibold'}, e('span',{className:'mr-2'}, showCOGSDetail ? '▼' : '▶'), '(-) Total COGS')
      ];
      if(showBudgetCols){
        cogsCells.push(e('td',{key:'budget',className:'py-2 px-3 text-right font-semibold'}, fmtCompact(cogsBudget || 0)));
        cogsCells.push(e('td',{key:'forecast',className:'py-2 px-3 text-right font-semibold'}, fmtCompact(cogsForecast || 0)));
      }
      cogsCells.push(e('td',{key:'actual',className:'py-2 px-3 text-right font-semibold'}, fmtCompact(cogsActual)));
      if(hasPYData && selYear === 2026){
        cogsCells.push(e('td',{key:'py',className:'py-2 px-3 text-right text-slate-400 font-semibold'}, fmtCompact(cogsPY || 0)));
      }
      cogsCells.push(e('td',{key:'var',className:'py-2 px-3 text-right font-semibold ' + cogsVarColor}, cogsVar !== null ? (cogsVar >= 0 ? '+' : '') + fmtCompact(cogsVar) : '—'));
      cogsCells.push(e('td',{key:'varpct',className:'py-2 px-3 text-right font-semibold ' + cogsVarColor}, cogsVar !== null && cogsComp ? fmtPercent(cogsVar / cogsComp) : '—'));
      
      tableRows.push(
        e('tr',{key:'cogs-header', className:'border-t border-slate-700/30 bg-slate-700/30 cursor-pointer hover:bg-slate-700/50', onClick: function(){ setShowCOGSDetail(!showCOGSDetail); }}, cogsCells)
      );
      if(showCOGSDetail){ COGS_ITEMS.forEach(function(item){ tableRows.push(PLRow(item, item, true, false, true, false)); }); }
      
      tableRows.push(PLRow('Gross Margin', 'Gross Margin', false, true, false, false));
      tableRows.push(PLRow('GM %', 'GM %', true, false, false, true));
      
      // OPEX expandable header
      var opexActual = getValue(fin.actual, 'OPEX');
      var opexBudget = showBudgetCols ? getBudgetValue('OPEX') : null;
      var opexForecast = showBudgetCols ? getForecastValue('OPEX') : null;
      var opexPY = hasPYData ? getPYValue('OPEX') : null;
      var opexComp = showBudgetCols ? opexBudget : opexPY;
      var opexVar = opexComp !== null ? opexActual - opexComp : null;
      var opexVarColor = opexVar !== null ? (opexVar <= 0 ? 'text-emerald-400' : 'text-rose-400') : '';
      
      var opexCells = [
        e('td',{key:'label',className:'py-2 pl-2 font-semibold'}, e('span',{className:'mr-2'}, showOPEXDetail ? '▼' : '▶'), '(-) Total OPEX')
      ];
      if(showBudgetCols){
        opexCells.push(e('td',{key:'budget',className:'py-2 px-3 text-right font-semibold'}, fmtCompact(opexBudget || 0)));
        opexCells.push(e('td',{key:'forecast',className:'py-2 px-3 text-right font-semibold'}, fmtCompact(opexForecast || 0)));
      }
      opexCells.push(e('td',{key:'actual',className:'py-2 px-3 text-right font-semibold'}, fmtCompact(opexActual)));
      if(hasPYData && selYear === 2026){
        opexCells.push(e('td',{key:'py',className:'py-2 px-3 text-right text-slate-400 font-semibold'}, fmtCompact(opexPY || 0)));
      }
      opexCells.push(e('td',{key:'var',className:'py-2 px-3 text-right font-semibold ' + opexVarColor}, opexVar !== null ? (opexVar >= 0 ? '+' : '') + fmtCompact(opexVar) : '—'));
      opexCells.push(e('td',{key:'varpct',className:'py-2 px-3 text-right font-semibold ' + opexVarColor}, opexVar !== null && opexComp ? fmtPercent(opexVar / opexComp) : '—'));
      
      tableRows.push(
        e('tr',{key:'opex-header', className:'border-t border-slate-700/30 bg-slate-700/30 cursor-pointer hover:bg-slate-700/50', onClick: function(){ setShowOPEXDetail(!showOPEXDetail); }}, opexCells)
      );
      if(showOPEXDetail){ OPEX_ITEMS.forEach(function(item){ tableRows.push(PLRow(item, item, true, false, true, false)); }); }
      
      // EBITDA row
      var ebitdaActual = getValue(fin.actual, 'EBITDA');
      var ebitdaBudget = showBudgetCols ? getBudgetValue('EBITDA') : null;
      var ebitdaForecast = showBudgetCols ? getForecastValue('EBITDA') : null;
      var ebitdaPY = hasPYData ? getPYValue('EBITDA') : null;
      var ebitdaComp = showBudgetCols ? ebitdaBudget : ebitdaPY;
      var ebitdaVar = ebitdaComp !== null ? ebitdaActual - ebitdaComp : null;
      var ebitdaVarColor = ebitdaVar !== null ? (ebitdaVar >= 0 ? 'text-emerald-400' : 'text-rose-400') : '';
      
      var ebitdaCells = [
        e('td',{key:'label',className:'py-3 pl-2 font-bold text-blue-300'},'EBITDA')
      ];
      if(showBudgetCols){
        ebitdaCells.push(e('td',{key:'budget',className:'py-3 px-3 text-right font-bold'}, fmtCompact(ebitdaBudget || 0)));
        ebitdaCells.push(e('td',{key:'forecast',className:'py-3 px-3 text-right font-bold'}, fmtCompact(ebitdaForecast || 0)));
      }
      ebitdaCells.push(e('td',{key:'actual',className:'py-3 px-3 text-right font-bold text-blue-300'}, fmtCompact(ebitdaActual)));
      if(hasPYData && selYear === 2026){
        ebitdaCells.push(e('td',{key:'py',className:'py-3 px-3 text-right text-slate-400 font-bold'}, fmtCompact(ebitdaPY || 0)));
      }
      ebitdaCells.push(e('td',{key:'var',className:'py-3 px-3 text-right font-bold ' + ebitdaVarColor}, ebitdaVar !== null ? (ebitdaVar >= 0 ? '+' : '') + fmtCompact(ebitdaVar) : '—'));
      ebitdaCells.push(e('td',{key:'varpct',className:'py-3 px-3 text-right font-bold ' + ebitdaVarColor}, ebitdaVar !== null && ebitdaComp ? fmtPercent(ebitdaVar / ebitdaComp) : '—'));
      
      tableRows.push(e('tr',{key:'ebitda', className:'border-t-2 border-blue-500 bg-blue-900/20'}, ebitdaCells));
      tableRows.push(PLRow('EBITDA %', 'EBTIDA %', true, false, false, true));
      
      // Build header row
      var headerCells = [e('th',{key:'label',className:'text-left py-2 px-2'},'Line Item')];
      if(showBudgetCols){
        headerCells.push(e('th',{key:'budget',className:'text-right py-2 px-3'},'Budget'));
        headerCells.push(e('th',{key:'forecast',className:'text-right py-2 px-3'},'Forecast'));
      }
      headerCells.push(e('th',{key:'actual',className:'text-right py-2 px-3'},'Actual'));
      if(hasPYData && selYear === 2026){
        headerCells.push(e('th',{key:'py',className:'text-right py-2 px-3'},'PY'));
      }
      headerCells.push(e('th',{key:'var',className:'text-right py-2 px-3'}, compLabel + ' $'));
      headerCells.push(e('th',{key:'varpct',className:'text-right py-2 px-3'}, compLabel + ' %'));
      
      var subtitle = showBudgetCols 
        ? 'Click COGS/OPEX to expand. Costs: green = under budget, red = over.'
        : 'Comparing to Prior Year. Click COGS/OPEX to expand.';
      
      return e('div',{className:'card p-6'},
        e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'P&L Summary'),
        e('div',{className:'text-xs text-slate-400 mb-3'}, subtitle),
        e('div',{className:'overflow-x-auto'},
          e('table',{className:'w-full text-sm'},
            e('thead',null,
              e('tr',{className:'text-xs text-slate-400 border-b border-slate-600'}, headerCells)
            ),
            e('tbody',null, tableRows)
          )
        )
      );
    }

    function TrendCharts(){
      var revenueChartId = 'fin-revenue-chart-' + selectedMonth;
      var ebitdaChartId = 'fin-ebitda-chart-' + selectedMonth;
      // Calculate year explicitly to avoid any closure issues
      var displayYear = parseMonthKey(selectedMonth) ? parseMonthKey(selectedMonth).year : 2025;
      
      React.useEffect(function(){
        var monthLabels = [];
        var actualRevData = [], budgetRevData = [], forecastRevData = [];
        var actualEbitdaData = [], budgetEbitdaData = [], forecastEbitdaData = [];
        var monthNames = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        
        for(var m = 1; m <= 12; m++){
          monthLabels.push(monthNames[m]);
          var actKey = m + '/' + displayYear;
          var budKey = m + '/2026';
          var actRow = getRow(fin.actual, actKey);
          var budRow = getRow(fin.budget, budKey);
          var fcRow = getRow(fin.forecast, budKey);
          
          if(fin.actualMonths.indexOf(actKey) >= 0){
            actualRevData.push(actRow ? actRow['Total Rev'] : 0);
            actualEbitdaData.push(actRow ? actRow['EBITDA'] : 0);
          } else {
            actualRevData.push(null);
            actualEbitdaData.push(null);
          }
          budgetRevData.push(budRow ? budRow['Total Rev'] : 0);
          forecastRevData.push(fcRow ? fcRow['Total Rev'] : 0);
          budgetEbitdaData.push(budRow ? budRow['EBITDA'] : 0);
          forecastEbitdaData.push(fcRow ? fcRow['EBITDA'] : 0);
        }
        
        createWhenReady(revenueChartId, function(ctx){
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: monthLabels,
              datasets: [
                { label: 'Budget', data: budgetRevData, backgroundColor: 'rgba(148,163,184,0.4)', borderColor: '#94a3b8', borderWidth: 1 },
                { label: 'Forecast', data: forecastRevData, backgroundColor: 'rgba(234,179,8,0.4)', borderColor: '#eab308', borderWidth: 1 },
                { label: 'Actual', data: actualRevData, backgroundColor: 'rgba(59,130,246,0.8)', borderColor: '#3b82f6', borderWidth: 1 }
              ]
            },
            options: {
              responsive: false, maintainAspectRatio: false, animation: false,
              legend: { display: true, position: 'top', labels: { fontColor: '#9fb0c7' } },
              tooltips: { mode: 'index', intersect: false, callbacks: { label: function(t,d){ return d.datasets[t.datasetIndex].label + ': ' + fmtCompact(t.yLabel); } } },
              scales: { xAxes: [{ ticks: { fontColor: '#9fb0c7' }, gridLines: gridLines() }], yAxes: [{ ticks: { fontColor: '#9fb0c7', callback: function(v){ return fmtCompact(v); } }, gridLines: gridLines() }] }
            }
          });
        });
        
        createWhenReady(ebitdaChartId, function(ctx){
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: monthLabels,
              datasets: [
                { label: 'Budget', data: budgetEbitdaData, backgroundColor: 'rgba(148,163,184,0.4)', borderColor: '#94a3b8', borderWidth: 1 },
                { label: 'Forecast', data: forecastEbitdaData, backgroundColor: 'rgba(234,179,8,0.4)', borderColor: '#eab308', borderWidth: 1 },
                { label: 'Actual', data: actualEbitdaData, backgroundColor: 'rgba(16,185,129,0.8)', borderColor: '#10b981', borderWidth: 1 }
              ]
            },
            options: {
              responsive: false, maintainAspectRatio: false, animation: false,
              legend: { display: true, position: 'top', labels: { fontColor: '#9fb0c7' } },
              tooltips: { mode: 'index', intersect: false, callbacks: { label: function(t,d){ return d.datasets[t.datasetIndex].label + ': ' + fmtCompact(t.yLabel); } } },
              scales: { xAxes: [{ ticks: { fontColor: '#9fb0c7' }, gridLines: gridLines() }], yAxes: [{ ticks: { fontColor: '#9fb0c7', callback: function(v){ return fmtCompact(v); } }, gridLines: gridLines() }] }
            }
          });
        });
      }, [displayYear, selectedMonth, fin.actual, fin.budget, fin.forecast]);
      
      return e('div',{className:'grid grid-cols-1 xl:grid-cols-2 gap-6'},
        e('div',{className:'card p-6'},
          e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'Revenue Trend (' + displayYear + ')'),
          e('div',{style:{height:'280px'}}, e('canvas',{id:revenueChartId}))
        ),
        e('div',{className:'card p-6'},
          e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'EBITDA Trend (' + displayYear + ')'),
          e('div',{style:{height:'280px'}}, e('canvas',{id:ebitdaChartId}))
        )
      );
    }
    
    // Revenue Mix - Cash vs Non-Cash breakdown
    function RevenueMix(){
      var totalRev = getValue(fin.actual, 'Total Rev');
      var cashRev = getValue(fin.actual, 'Cash');
      var premiereRev = getValue(fin.actual, 'Premiere');
      var barterRev = getValue(fin.actual, 'Other Barter');
      var miscRev = getValue(fin.actual, 'Misc.');
      var nonCashRev = premiereRev + barterRev;
      
      var cashPct = totalRev ? cashRev / totalRev : 0;
      var nonCashPct = totalRev ? nonCashRev / totalRev : 0;
      var premierePct = totalRev ? premiereRev / totalRev : 0;
      var barterPct = totalRev ? barterRev / totalRev : 0;
      var miscPct = totalRev ? miscRev / totalRev : 0;
      
      // Get PY comparison
      var pyTotalRev = getPYValue('Total Rev');
      var pyCashRev = getPYValue('Cash');
      var pyNonCashRev = getPYValue('Premiere') + getPYValue('Other Barter');
      var pyCashPct = pyTotalRev ? pyCashRev / pyTotalRev : 0;
      var pyNonCashPct = pyTotalRev ? pyNonCashRev / pyTotalRev : 0;
      
      return e('div',{className:'card p-6'},
        e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'Revenue Mix'),
        e('div',{className:'grid grid-cols-2 md:grid-cols-5 gap-4'},
          e('div',{className:'bg-slate-700/30 rounded-lg p-4'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Cash Revenue'),
            e('div',{className:'text-xl font-bold text-emerald-400'}, fmtPercent(cashPct)),
            e('div',{className:'text-sm text-slate-300'}, fmtCompact(cashRev)),
            hasPYData ? e('div',{className:'text-xs text-slate-500 mt-1'}, 'PY: ' + fmtPercent(pyCashPct)) : null
          ),
          e('div',{className:'bg-slate-700/30 rounded-lg p-4'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Non-Cash Revenue'),
            e('div',{className:'text-xl font-bold text-blue-400'}, fmtPercent(nonCashPct)),
            e('div',{className:'text-sm text-slate-300'}, fmtCompact(nonCashRev)),
            hasPYData ? e('div',{className:'text-xs text-slate-500 mt-1'}, 'PY: ' + fmtPercent(pyNonCashPct)) : null
          ),
          e('div',{className:'bg-slate-700/30 rounded-lg p-4'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Premiere'),
            e('div',{className:'text-xl font-bold text-violet-400'}, fmtPercent(premierePct)),
            e('div',{className:'text-sm text-slate-300'}, fmtCompact(premiereRev))
          ),
          e('div',{className:'bg-slate-700/30 rounded-lg p-4'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Other Barter'),
            e('div',{className:'text-xl font-bold text-amber-400'}, fmtPercent(barterPct)),
            e('div',{className:'text-sm text-slate-300'}, fmtCompact(barterRev))
          ),
          e('div',{className:'bg-slate-700/30 rounded-lg p-4'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Misc.'),
            e('div',{className:'text-xl font-bold text-pink-400'}, fmtPercent(miscPct)),
            e('div',{className:'text-sm text-slate-300'}, fmtCompact(miscRev))
          )
        ),
        e('div',{className:'mt-4'},
          e('div',{className:'text-xs text-slate-400 mb-2'},'Revenue Composition'),
          e('div',{className:'h-4 rounded-full overflow-hidden flex'},
            e('div',{className:'bg-emerald-500', style:{width: fmtPercent(cashPct)}}),
            e('div',{className:'bg-violet-500', style:{width: fmtPercent(premierePct)}}),
            e('div',{className:'bg-amber-500', style:{width: fmtPercent(barterPct)}}),
            e('div',{className:'bg-pink-500', style:{width: fmtPercent(Math.abs(miscPct))}})
          ),
          e('div',{className:'flex justify-between text-xs text-slate-500 mt-1'},
            e('span',null,'Cash ' + fmtPercent(cashPct)),
            e('span',null,'Premiere ' + fmtPercent(premierePct)),
            e('span',null,'Barter ' + fmtPercent(barterPct)),
            e('span',null,'Misc ' + fmtPercent(miscPct))
          )
        )
      );
    }

    function Controls(){
      var viewLabel = viewMode === 'monthly' ? 'MTD' : viewMode === 'ytd' ? 'YTD' : 'LTM';
      return e('div',{className:'card p-4 flex flex-wrap items-center gap-4 justify-between'},
        e('div',{className:'flex items-center gap-3'},
          e('h2',{className:'text-lg font-semibold'},'Financials'),
          e('span',{className:'pill'}, viewLabel + ' through ' + formatMonthKey(selectedMonth))
        ),
        e('div',{className:'flex items-center gap-3'},
          e('div',{className:'flex rounded-lg overflow-hidden border border-slate-600'},
            e('button',{className:'px-3 py-1.5 text-sm ' + (viewMode === 'monthly' ? 'bg-blue-600 text-white' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'), onClick: function(){ setViewMode('monthly'); }},'MTD'),
            e('button',{className:'px-3 py-1.5 text-sm ' + (viewMode === 'ytd' ? 'bg-blue-600 text-white' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'), onClick: function(){ setViewMode('ytd'); }},'YTD'),
            e('button',{className:'px-3 py-1.5 text-sm ' + (viewMode === 'ltm' ? 'bg-blue-600 text-white' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'), onClick: function(){ setViewMode('ltm'); }},'LTM')
          ),
          e('select',{className:'bg-slate-700 border border-slate-600 rounded-md px-3 py-1.5 text-sm', value: selectedMonth, onChange: function(ev){ setSelectedMonth(ev.target.value); }},
            fin.actualMonths.map(function(mk){ return e('option',{key:mk, value:mk}, formatMonthKey(mk)); })
          )
        )
      );
    }

    return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800'},
      e('div',{className:'max-w-[2000px] mx-auto px-4 lg:px-6 xl:px-8 py-6 space-y-6'},
        e(Controls),
        e(SummaryCards),
        e(RevenueMix),
        e(PLTable),
        e(TrendCharts),
        e('footer',{className:'text-xs text-slate-500 py-4 text-center flex justify-center gap-4'},
          e('span',null, 'Last actual: ' + formatMonthKey(fin.lastActualMonth)),
          e('span',null, '•'),
          e('span',null, 'Source: ' + (state.dataSource || 'Local')),
          e('span',null, '•'),
          e('span',null, 'Loaded: ' + (state.loadTime || '-'))
        )
      )
    );
  }

  /* ------------------------ Pacing Screen (Premiere Pacing) ------------------------ */
  function PacingScreen(){
    var _data = React.useState([]);
    var pacingData = _data[0], setPacingData = _data[1];
    
    var _budgetData = React.useState({});
    var budgetData = _budgetData[0], setBudgetData = _budgetData[1];
    
    var _err = React.useState(null);
    var err = _err[0], setErr = _err[1];
    
    var _selectedWeek = React.useState(null);
    var selectedWeek = _selectedWeek[0], setSelectedWeek = _selectedWeek[1];

    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var monthMap = {'1':'Jan','2':'Feb','3':'Mar','4':'Apr','5':'May','6':'Jun','7':'Jul','8':'Aug','9':'Sep','10':'Oct','11':'Nov','12':'Dec'};

    function getQuarterMonths(quarter) {
      var quarters = { 'Q1': ['Jan', 'Feb', 'Mar'], 'Q2': ['Apr', 'May', 'Jun'], 'Q3': ['Jul', 'Aug', 'Sep'], 'Q4': ['Oct', 'Nov', 'Dec'] };
      return quarters[quarter] || [];
    }

    function calculateQuarterTotal(rowData, quarter) {
      var quarterMonths = getQuarterMonths(quarter);
      return quarterMonths.reduce(function(sum, month) { return sum + (rowData[month] || 0); }, 0);
    }

    function calculateFYTotal(rowData) {
      if(!rowData) return 0;
      return months.reduce(function(sum, month) { return sum + (rowData[month] || 0); }, 0);
    }

    function findPriorWeek(currentWeek) {
      var currentIndex = pacingData.findIndex(function(row){ return row.dateStr === currentWeek.dateStr; });
      return currentIndex < pacingData.length - 1 ? pacingData[currentIndex + 1] : null;
    }

    function findPriorYearWeek(currentWeek) {
      var targetYear = currentWeek.year - 1;
      var currentMonth = currentWeek.date.getMonth();
      var currentDay = currentWeek.date.getDate();
      var bestMatch = null, bestDiff = Infinity;
      
      pacingData.forEach(function(row){
        if(row.year === targetYear){
          var rowMonth = row.date.getMonth();
          var rowDay = row.date.getDate();
          if(rowMonth === currentMonth){
            var dayDiff = Math.abs(rowDay - currentDay);
            if(dayDiff < bestDiff && dayDiff <= 7){ bestDiff = dayDiff; bestMatch = row; }
          }
        }
      });
      return bestMatch;
    }

    // Load pacing and budget data
    React.useEffect(function(){
      // Cache-buster to force fresh data on each load
      var cacheBuster = '&_t=' + Date.now();
      
      // Google Sheets published URL (primary), local files (fallback)
      var pacingPaths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSICXr4bYp-4kszCotiAIHXGBJElALUawJZBRit5fvs2fRD9DQZZSTN9lusoAkdazXDHKvjjh2n1sbB/pub?gid=415023472&single=true&output=csv' + cacheBuster,
        '/public/Pacing.csv','public/Pacing.csv','/Pacing.csv','Pacing.csv'
      ];
      var budgetPaths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSICXr4bYp-4kszCotiAIHXGBJElALUawJZBRit5fvs2fRD9DQZZSTN9lusoAkdazXDHKvjjh2n1sbB/pub?gid=586410538&single=true&output=csv' + cacheBuster,
        '/public/budget.csv','public/budget.csv','/budget.csv','budget.csv'
      ];
      
      Promise.all([
        loadCsvText(pacingPaths),
        loadCsvText(budgetPaths).catch(function(){ return null; })
      ]).then(function(results){
        var pacingTxt = results[0];
        var budgetTxt = results[1];
        
        // Process pacing data
        var parsed = Papa.parse(pacingTxt, {header: true, skipEmptyLines: true});
        var rawData = parsed.data || [];
        
        var processedData = rawData.map(function(row){
          var dateStr = row.DATE;
          var parsedDate = new Date(dateStr);
          if(isNaN(parsedDate.getTime())) return null;
          
          var processedRow = { date: parsedDate, dateStr: dateStr, year: parsedDate.getFullYear() };
          months.forEach(function(month){
            var value = row[month];
            if(value && value !== ''){
              var cleanValue = String(value).replace(/[$,]/g,'');
              var numValue = parseFloat(cleanValue);
              processedRow[month] = isNaN(numValue) ? 0 : numValue;
            } else {
              processedRow[month] = 0;
            }
          });
          return processedRow;
        }).filter(function(row){ return row !== null; });
        
        processedData.sort(function(a, b){ return b.date - a.date; });
        setPacingData(processedData);
        
        // Process budget data - organize by year
        // Format could be:
        // - Column-based: first col empty with month (1/2026), columns are line items (Premiere, Cash, etc.)
        // - Row-based: first col is "Line Item", columns are months (1/2026, 2/2026, etc.)
        if(budgetTxt){
          var budgetParsed = Papa.parse(budgetTxt, {header: true, skipEmptyLines: true});
          var budgetRows = budgetParsed.data || [];
          var budgetByYear = {};
          
          var cols = Object.keys(budgetRows[0] || {});
          console.log('Budget CSV columns:', cols);
          
          // Detect format: if first column is empty or matches month pattern, it's column-based
          var firstCol = cols[0] || '';
          var firstCellValue = (budgetRows[0] && budgetRows[0][firstCol]) ? budgetRows[0][firstCol].toString() : '';
          var isColumnBased = firstCol === '' || /^\d{1,2}\/\d{4}$/.test(firstCellValue);
          
          console.log('Budget format detected:', isColumnBased ? 'column-based' : 'row-based');
          
          if(isColumnBased){
            // Column-based: each row is a month, Premiere is a column
            budgetRows.forEach(function(row){
              var monthYearStr = (row[''] || row[firstCol] || '').toString().trim();
              var match = monthYearStr.match(/^(\d{1,2})\/(\d{4})$/);
              if(match){
                var monthNum = match[1];
                var year = parseInt(match[2], 10);
                var monthName = monthMap[monthNum];
                var premiereValue = parseFloat(String(row.Premiere || row['Premiere'] || '0').replace(/[$,]/g,'')) || 0;
                
                if(!budgetByYear[year]) budgetByYear[year] = {};
                budgetByYear[year][monthName] = premiereValue;
              }
            });
          } else {
            // Row-based: each row is a line item, months are columns
            var premiereRow = budgetRows.find(function(row){
              var lineItem = (row[firstCol] || row['Line Item'] || '').toString().trim();
              return lineItem.toLowerCase() === 'premiere';
            });
            
            if(premiereRow){
              cols.forEach(function(colName){
                if(colName === firstCol || colName === 'Line Item' || colName === '') return;
                var match = colName.match(/^(\d{1,2})\/(\d{4})$/);
                if(match){
                  var monthNum = match[1];
                  var year = parseInt(match[2], 10);
                  var monthName = monthMap[monthNum];
                  var premiereValue = parseFloat(String(premiereRow[colName] || '0').replace(/[$,]/g,'')) || 0;
                  
                  if(!budgetByYear[year]) budgetByYear[year] = {};
                  budgetByYear[year][monthName] = premiereValue;
                }
              });
            }
          }
          
          console.log('Budget by year:', budgetByYear);
          setBudgetData(budgetByYear);
        } else {
          console.log('No budget text loaded');
        }
        
        setErr(null);
        
        // Auto-select most recent week with ANY month > 0
        if(processedData.length > 0){
          var mostRecentWithData = null;
          for(var i = 0; i < processedData.length; i++){
            var row = processedData[i];
            var hasData = months.some(function(m){ return row[m] && row[m] > 0; });
            if(hasData){ mostRecentWithData = row; break; }
          }
          setSelectedWeek(mostRecentWithData || processedData[0]);
        }
      }).catch(function(error){
        setErr('Failed to load Pacing.csv. ' + (error.message || ''));
      });
    }, []);

    if(err){
      return e('main',{className:'space-y-4'},
        e('section',{className:'card p-6'},
          e('div',{className:'text-xl font-semibold mb-2 text-red-400'},'Premiere Pacing - Error'),
          e('div',{className:'text-slate-400'},err)
        )
      );
    }

    if(!pacingData.length || !selectedWeek){
      return e('main',{className:'space-y-4'},
        e('section',{className:'card p-6'},
          e('div',{className:'text-xl font-semibold mb-2'},'Premiere Pacing'),
          e('div',{className:'text-slate-400'},'Loading Pacing.csv...')
        )
      );
    }

    var priorWeek = findPriorWeek(selectedWeek);
    var priorYearWeek = findPriorYearWeek(selectedWeek);
    var currentYearBudget = budgetData[selectedWeek.year] || {};

    function PacingControls(){
      var availableWeeks = pacingData.filter(function(w){
        return months.some(function(m){ return w[m] && w[m] > 0; });
      }).slice(0, 50);
      
      return e('section',{className:'card p-4 flex flex-wrap items-center gap-3 justify-between'},
        e('div',{className:'flex items-center gap-2'},[
          e('div',{key:'t',className:'text-base font-semibold'},'Premiere Pacing'),
          e('span',{key:'p',className:'pill'},'Week of ' + selectedWeek.dateStr)
        ]),
        e('div',{className:'flex items-center gap-2'},[
          e('select',{
            key:'s',
            className:'bg-[#0e141c] border border-slate-700 rounded-md px-3 py-2 text-sm',
            value: selectedWeek.dateStr,
            onChange: function(ev){
              var selected = pacingData.find(function(row){ return row.dateStr === ev.target.value; });
              if(selected) setSelectedWeek(selected);
            }
          }, availableWeeks.map(function(week){
            return e('option',{key: week.dateStr, value: week.dateStr}, 'Week of ' + week.dateStr + ' (' + week.year + ')');
          }))
        ])
      );
    }

    function SummaryCards(){
      var currentYearTotal = calculateFYTotal(selectedWeek);
      var priorWeekTotal = priorWeek ? calculateFYTotal(priorWeek) : 0;
      var priorYearTotal = priorYearWeek ? calculateFYTotal(priorYearWeek) : 0;
      var budgetTotal = calculateFYTotal(currentYearBudget);
      
      var wowDollarChange = priorWeekTotal ? currentYearTotal - priorWeekTotal : null;
      var wowPctChange = priorWeekTotal ? wowDollarChange / priorWeekTotal : null;
      var yoyDollarChange = priorYearTotal ? currentYearTotal - priorYearTotal : null;
      var yoyPctChange = priorYearTotal ? yoyDollarChange / priorYearTotal : null;
      var budgetVariance = budgetTotal ? currentYearTotal - budgetTotal : null;
      var budgetAttainment = budgetTotal ? currentYearTotal / budgetTotal : null;

      function SummaryCard(props){
        return e('div',{className:'card p-4 ' + (props.highlight ? 'ring-2 ring-blue-500' : '')},
          e('div',{className:'text-xs font-medium text-slate-400 uppercase tracking-wider mb-2'}, props.title),
          e('div',{className:'text-2xl font-bold ' + (props.valueClass || (props.highlight ? 'text-blue-300' : 'text-slate-100')) + ' mb-1'}, props.value),
          props.subtitle ? e('div',{className:'text-xs text-slate-500 mb-2'}, props.subtitle) : null,
          props.changes ? e('div',{className:'text-xs space-y-1'}, props.changes) : null
        );
      }

      return e('div',{className:'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4'}, [
        e(SummaryCard,{ 
          key:'curr', 
          title: 'Current YTD Pacing', 
          value: fmtCompact(currentYearTotal), 
          subtitle: 'Week of ' + selectedWeek.dateStr, 
          highlight: true 
        }),
        e(SummaryCard,{ 
          key:'pw', 
          title: 'Prior Week YTD', 
          value: fmtCompact(priorWeekTotal), 
          subtitle: priorWeek ? 'Week of ' + priorWeek.dateStr : 'N/A',
          changes: priorWeek ? [
            e('div',{key:'d',className: wowDollarChange >= 0 ? 'text-emerald-400' : 'text-rose-400'}, 
              'WoW: ' + (wowDollarChange >= 0 ? '+' : '') + fmtCompact(wowDollarChange) + ' (' + fmtPercent(wowPctChange) + ')')
          ] : null
        }),
        e(SummaryCard,{ 
          key:'py', 
          title: 'Prior Year (Same Week)', 
          value: fmtCompact(priorYearTotal), 
          subtitle: priorYearWeek ? 'Week of ' + priorYearWeek.dateStr : 'N/A',
          changes: priorYearWeek ? [
            e('div',{key:'d',className: yoyDollarChange >= 0 ? 'text-emerald-400' : 'text-rose-400'}, 
              'YoY: ' + (yoyDollarChange >= 0 ? '+' : '') + fmtCompact(yoyDollarChange) + ' (' + fmtPercent(yoyPctChange) + ')')
          ] : null
        }),
        e(SummaryCard,{ 
          key:'budget', 
          title: 'Budget Attainment', 
          value: budgetTotal ? fmtPercent(budgetAttainment) : 'N/A',
          valueClass: budgetAttainment >= 1 ? 'text-emerald-400' : budgetAttainment >= 0.9 ? 'text-yellow-400' : 'text-rose-400',
          subtitle: budgetTotal ? fmtCompact(budgetTotal) + ' annual budget' : 'No budget data',
          changes: budgetTotal ? [
            e('div',{key:'v',className: budgetVariance >= 0 ? 'text-emerald-400' : 'text-rose-400'}, 
              (budgetVariance >= 0 ? '+' : '') + fmtCompact(budgetVariance) + ' vs budget')
          ] : null
        })
      ]);
    }

    function ComparisonChart(){
      var chartId = 'premiere-comparison-chart';
      
      React.useEffect(function(){
        if(!selectedWeek) return;
        
        var currentData = months.map(function(m){ return selectedWeek[m] || 0; });
        var priorYearData = months.map(function(m){ return priorYearWeek ? (priorYearWeek[m] || 0) : 0; });
        
        createWhenReady(chartId, function(ctx){
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: months,
              datasets: [
                {
                  label: selectedWeek.year + ' Pacing',
                  data: currentData,
                  backgroundColor: 'rgba(59, 130, 246, 0.8)',
                  borderColor: '#3b82f6',
                  borderWidth: 1
                },
                {
                  label: (selectedWeek.year - 1) + ' (Same Week)',
                  data: priorYearData,
                  backgroundColor: 'rgba(148, 163, 184, 0.5)',
                  borderColor: '#94a3b8',
                  borderWidth: 1
                }
              ]
            },
            options: {
              responsive: false,
              maintainAspectRatio: false,
              animation: false,
              legend: { display: true, position: 'top', labels: { fontColor: '#9fb0c7' } },
              tooltips: { 
                mode: 'index', 
                intersect: false,
                callbacks: {
                  label: function(tooltipItem, data) {
                    var label = data.datasets[tooltipItem.datasetIndex].label || '';
                    return label + ': ' + fmtCompact(tooltipItem.yLabel);
                  }
                }
              },
              scales: {
                xAxes: [{ ticks: { fontColor: '#9fb0c7' }, gridLines: gridLines() }],
                yAxes: [{ ticks: { fontColor: '#9fb0c7', callback: function(v){ return fmtCompact(v); } }, gridLines: gridLines() }]
              }
            }
          });
        });
      }, [selectedWeek, priorYearWeek]);
      
      return e('div',{className:'card p-6'},
        e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'Current Year vs Prior Year by Month'),
        e('div',{style:{height:'320px'}},
          e('canvas',{id:chartId})
        )
      );
    }

    function BudgetTracker(){
      var budgetTotal = calculateFYTotal(currentYearBudget);
      var currentTotal = calculateFYTotal(selectedWeek);
      var variance = currentTotal - budgetTotal;
      var variancePct = budgetTotal ? variance / budgetTotal : null;
      var attainmentPct = budgetTotal ? currentTotal / budgetTotal : null;
      
      if(!budgetTotal){
        return e('div',{className:'card p-6'},
          e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'Pacing vs Budget'),
          e('div',{className:'text-slate-400'},'No budget data available for ' + selectedWeek.year)
        );
      }
      
      return e('div',{className:'card p-6'},
        e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'Pacing vs Budget (' + selectedWeek.year + ')'),
        e('div',{className:'grid grid-cols-1 md:grid-cols-3 gap-4 mb-6'},[
          e('div',{key:'budget',className:'bg-slate-800/50 rounded-lg p-4 text-center'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Annual Budget'),
            e('div',{className:'text-xl font-bold text-slate-100'},fmtCompact(budgetTotal))
          ),
          e('div',{key:'current',className:'bg-slate-800/50 rounded-lg p-4 text-center'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Current Pacing'),
            e('div',{className:'text-xl font-bold text-blue-400'},fmtCompact(currentTotal))
          ),
          e('div',{key:'variance',className:'bg-slate-800/50 rounded-lg p-4 text-center'},
            e('div',{className:'text-xs text-slate-400 uppercase mb-1'},'Variance'),
            e('div',{className:'text-xl font-bold ' + (variance >= 0 ? 'text-emerald-400' : 'text-rose-400')},
              (variance >= 0 ? '+' : '') + fmtCompact(variance) + ' (' + fmtPercent(variancePct) + ')')
          )
        ]),
        e('div',{className:'mb-2 flex justify-between text-sm'},
          e('span',{className:'text-slate-400'},'Budget Attainment'),
          e('span',{className:'font-medium ' + (attainmentPct >= 1 ? 'text-emerald-400' : attainmentPct >= 0.9 ? 'text-yellow-400' : 'text-rose-400')},
            fmtPercent(attainmentPct))
        ),
        e('div',{className:'w-full bg-slate-700 rounded-full h-4 overflow-hidden'},
          e('div',{className:'h-full rounded-full transition-all ' + (attainmentPct >= 1 ? 'bg-emerald-500' : attainmentPct >= 0.9 ? 'bg-yellow-500' : 'bg-blue-500'),
            style:{width: Math.min(100, (attainmentPct || 0) * 100) + '%'}})
        ),
        e('div',{className:'mt-4 overflow-x-auto'},
          e('table',{className:'w-full text-sm'},
            e('thead',null,
              e('tr',{className:'text-xs text-slate-400 border-b border-slate-600'},
                e('th',{className:'text-left py-2 px-2'},'Month'),
                e('th',{className:'text-right py-2 px-2'},'Budget'),
                e('th',{className:'text-right py-2 px-2'},'Pacing'),
                e('th',{className:'text-right py-2 px-2'},'Variance'),
                e('th',{className:'text-right py-2 px-2'},'% of Budget')
              )
            ),
            e('tbody',null,
              months.map(function(month){
                var budget = currentYearBudget[month] || 0;
                var pacing = selectedWeek[month] || 0;
                var monthVar = pacing - budget;
                var monthPct = budget ? pacing / budget : null;
                
                return e('tr',{key:month, className:'border-t border-slate-700/30'},
                  e('td',{className:'py-2 px-2 font-medium'},month),
                  e('td',{className:'py-2 px-2 text-right'},fmtCompact(budget)),
                  e('td',{className:'py-2 px-2 text-right'},fmtCompact(pacing)),
                  e('td',{className:'py-2 px-2 text-right ' + (monthVar >= 0 ? 'text-emerald-400' : 'text-rose-400')},
                    (monthVar >= 0 ? '+' : '') + fmtCompact(monthVar)),
                  e('td',{className:'py-2 px-2 text-right ' + (monthPct >= 1 ? 'text-emerald-400' : monthPct >= 0.9 ? 'text-yellow-400' : 'text-slate-300')},
                    monthPct !== null ? fmtPercent(monthPct) : '—')
                );
              }),
              e('tr',{key:'total', className:'border-t-2 border-slate-600 font-bold bg-slate-700/20'},
                e('td',{className:'py-2 px-2'},'FY TOTAL'),
                e('td',{className:'py-2 px-2 text-right'},fmtCompact(budgetTotal)),
                e('td',{className:'py-2 px-2 text-right'},fmtCompact(currentTotal)),
                e('td',{className:'py-2 px-2 text-right ' + (variance >= 0 ? 'text-emerald-400' : 'text-rose-400')},
                  (variance >= 0 ? '+' : '') + fmtCompact(variance)),
                e('td',{className:'py-2 px-2 text-right ' + (attainmentPct >= 1 ? 'text-emerald-400' : attainmentPct >= 0.9 ? 'text-yellow-400' : 'text-slate-300')},
                  fmtPercent(attainmentPct))
              )
            )
          )
        )
      );
    }

    function DetailedTable(){
      var quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
      var tableRows = [];
      
      months.forEach(function(month){
        var current = selectedWeek[month] || 0;
        var budget = currentYearBudget[month] || 0;
        var priorWeekValue = priorWeek ? (priorWeek[month] || 0) : 0;
        var priorYearValue = priorYearWeek ? (priorYearWeek[month] || 0) : 0;
        
        var budgetVar = current - budget;
        var budgetVarPct = budget ? budgetVar / budget : null;
        var weekChange = current - priorWeekValue;
        var yearChange = current - priorYearValue;
        var yearChangePct = priorYearValue ? yearChange / priorYearValue : null;
        
        // Row background based on budget performance
        var rowBg = '';
        if(budget > 0){
          if(budgetVarPct >= 0) rowBg = 'bg-emerald-900/15';
          else if(budgetVarPct < -0.1) rowBg = 'bg-rose-900/20';
          else rowBg = 'bg-rose-900/10';
        }
        
        tableRows.push(
          e('tr',{key: month, className:'border-t border-slate-700/30 ' + rowBg},
            e('td',{key:'m',className:'py-2 px-2 font-medium'}, month),
            e('td',{key:'b',className:'py-2 px-2 text-right text-slate-400'}, budget ? fmtCompact(budget) : '—'),
            e('td',{key:'c',className:'py-2 px-2 text-right font-semibold'}, fmtCompact(current)),
            e('td',{key:'bv',className:'py-2 px-2 text-right ' + (budgetVar >= 0 ? 'text-emerald-400' : 'text-rose-400')}, 
              budget ? (budgetVar >= 0 ? '+' : '') + fmtCompact(budgetVar) : '—'),
            e('td',{key:'bvp',className:'py-2 px-2 text-right ' + (budgetVarPct >= 0 ? 'text-emerald-400' : 'text-rose-400')}, 
              budget ? fmtPercent(budgetVarPct) : '—'),
            e('td',{key:'py',className:'py-2 px-2 text-right text-slate-400'}, fmtCompact(priorYearValue)),
            e('td',{key:'yoyd',className:'py-2 px-2 text-right ' + clsDelta(yearChange, true)}, 
              priorYearValue ? (yearChange >= 0 ? '+' : '') + fmtCompact(yearChange) : '—'),
            e('td',{key:'yoyp',className:'py-2 px-2 text-right ' + clsDelta(yearChangePct, true)}, 
              priorYearWeek ? fmtPercent(yearChangePct) : '—'),
            e('td',{key:'wowd',className:'py-2 px-2 text-right ' + clsDelta(weekChange, true)}, 
              priorWeek ? (weekChange >= 0 ? '+' : '') + fmtCompact(weekChange) : '—')
          )
        );
      });
      
      // Quarterly rollups
      quarters.forEach(function(quarter){
        var current = calculateQuarterTotal(selectedWeek, quarter);
        var budget = calculateQuarterTotal(currentYearBudget, quarter);
        var priorWeekValue = priorWeek ? calculateQuarterTotal(priorWeek, quarter) : 0;
        var priorYearValue = priorYearWeek ? calculateQuarterTotal(priorYearWeek, quarter) : 0;
        
        var budgetVar = current - budget;
        var budgetVarPct = budget ? budgetVar / budget : null;
        var weekChange = current - priorWeekValue;
        var yearChange = current - priorYearValue;
        var yearChangePct = priorYearValue ? yearChange / priorYearValue : null;
        
        tableRows.push(
          e('tr',{key: quarter, className:'border-t-2 border-slate-600 font-semibold bg-slate-700/30'},
            e('td',{key:'m',className:'py-2 px-2'}, quarter + ' Total'),
            e('td',{key:'b',className:'py-2 px-2 text-right text-slate-400'}, budget ? fmtCompact(budget) : '—'),
            e('td',{key:'c',className:'py-2 px-2 text-right'}, fmtCompact(current)),
            e('td',{key:'bv',className:'py-2 px-2 text-right ' + (budgetVar >= 0 ? 'text-emerald-400' : 'text-rose-400')}, 
              budget ? (budgetVar >= 0 ? '+' : '') + fmtCompact(budgetVar) : '—'),
            e('td',{key:'bvp',className:'py-2 px-2 text-right ' + (budgetVarPct >= 0 ? 'text-emerald-400' : 'text-rose-400')}, 
              budget ? fmtPercent(budgetVarPct) : '—'),
            e('td',{key:'py',className:'py-2 px-2 text-right text-slate-400'}, fmtCompact(priorYearValue)),
            e('td',{key:'yoyd',className:'py-2 px-2 text-right ' + clsDelta(yearChange, true)}, 
              priorYearValue ? (yearChange >= 0 ? '+' : '') + fmtCompact(yearChange) : '—'),
            e('td',{key:'yoyp',className:'py-2 px-2 text-right ' + clsDelta(yearChangePct, true)}, 
              priorYearWeek ? fmtPercent(yearChangePct) : '—'),
            e('td',{key:'wowd',className:'py-2 px-2 text-right ' + clsDelta(weekChange, true)}, 
              priorWeek ? (weekChange >= 0 ? '+' : '') + fmtCompact(weekChange) : '—')
          )
        );
      });
      
      // FY Total
      var fyTotal = calculateFYTotal(selectedWeek);
      var fyBudget = calculateFYTotal(currentYearBudget);
      var fyPW = priorWeek ? calculateFYTotal(priorWeek) : 0;
      var fyPY = priorYearWeek ? calculateFYTotal(priorYearWeek) : 0;
      var fyBudgetVar = fyTotal - fyBudget;
      var fyBudgetVarPct = fyBudget ? fyBudgetVar / fyBudget : null;
      var fyWeekChange = fyTotal - fyPW;
      var fyYearChange = fyTotal - fyPY;
      
      tableRows.push(
        e('tr',{key: 'fy', className:'border-t-2 border-blue-500 font-bold bg-blue-900/20'},
          e('td',{key:'m',className:'py-2 px-2'}, 'FY TOTAL'),
          e('td',{key:'b',className:'py-2 px-2 text-right text-slate-300'}, fyBudget ? fmtCompact(fyBudget) : '—'),
          e('td',{key:'c',className:'py-2 px-2 text-right text-blue-300'}, fmtCompact(fyTotal)),
          e('td',{key:'bv',className:'py-2 px-2 text-right ' + (fyBudgetVar >= 0 ? 'text-emerald-400' : 'text-rose-400')}, 
            fyBudget ? (fyBudgetVar >= 0 ? '+' : '') + fmtCompact(fyBudgetVar) : '—'),
          e('td',{key:'bvp',className:'py-2 px-2 text-right ' + (fyBudgetVarPct >= 0 ? 'text-emerald-400' : 'text-rose-400')}, 
            fyBudget ? fmtPercent(fyBudgetVarPct) : '—'),
          e('td',{key:'py',className:'py-2 px-2 text-right text-slate-400'}, fmtCompact(fyPY)),
          e('td',{key:'yoyd',className:'py-2 px-2 text-right ' + clsDelta(fyYearChange, true)}, 
            fyPY ? (fyYearChange >= 0 ? '+' : '') + fmtCompact(fyYearChange) : '—'),
          e('td',{key:'yoyp',className:'py-2 px-2 text-right ' + clsDelta(fyPY ? fyYearChange/fyPY : null, true)}, 
            fyPY ? fmtPercent(fyYearChange/fyPY) : '—'),
          e('td',{key:'wowd',className:'py-2 px-2 text-right ' + clsDelta(fyWeekChange, true)}, 
            fyPW ? (fyWeekChange >= 0 ? '+' : '') + fmtCompact(fyWeekChange) : '—')
        )
      );
      
      return e('div',{className:'card p-6'},
        e('h3',{className:'text-lg font-semibold mb-4 text-slate-200'},'Detailed Pacing Analysis'),
        e('div',{className:'text-xs text-slate-400 mb-3'},'Rows highlighted green (beating budget) or red (missing budget)'),
        e('div',{className:'overflow-x-auto'},
          e('table',{className:'w-full text-sm'},
            e('thead',null,
              e('tr',{className:'text-xs text-slate-400 border-b border-slate-600'},
                e('th',{key:'p',className:'text-left py-2 px-2'},'Period'),
                e('th',{key:'b',className:'text-right py-2 px-2'},'Budget'),
                e('th',{key:'c',className:'text-right py-2 px-2'},'Current'),
                e('th',{key:'bv',className:'text-right py-2 px-2'},'vs Budget $'),
                e('th',{key:'bvp',className:'text-right py-2 px-2'},'vs Budget %'),
                e('th',{key:'py',className:'text-right py-2 px-2'},(selectedWeek.year-1)+' PY'),
                e('th',{key:'yoyd',className:'text-right py-2 px-2'},'YoY $'),
                e('th',{key:'yoyp',className:'text-right py-2 px-2'},'YoY %'),
                e('th',{key:'wowd',className:'text-right py-2 px-2'},'WoW $')
              )
            ),
            e('tbody',null, tableRows)
          )
        )
      );
    }

    return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800'},
      e('div',{className:'max-w-[2000px] mx-auto px-4 lg:px-6 xl:px-8 py-6 space-y-6'},
        e(PacingControls),
        e(SummaryCards),
        e('div',{className:'grid grid-cols-1 xl:grid-cols-2 gap-6'},
          e(ComparisonChart),
          e(BudgetTracker)
        ),
        e(DetailedTable),
        e('footer',{className:'text-xs text-slate-500 py-4 text-center'},
          'Data as of ' + selectedWeek.dateStr
        )
      )
    );
  }

  /* ------------------------ Diversification Screen ------------------------ */
  function DiversificationScreen() {
    var _state = React.useState({ data: null, loading: true, error: null, errorDetails: null, refreshing: false });
    var state = _state[0], setState = _state[1];

    function loadData() {
      fetch('/api/diversification')
        .then(function(res) {
          return res.json().then(function(data) {
            if (!res.ok) {
              var err = new Error(data.message || 'Unknown error');
              err.details = data;
              throw err;
            }
            return data;
          });
        })
        .then(function(data) {
          setState({ data: data, loading: false, error: null, errorDetails: null, refreshing: false });
        })
        .catch(function(err) {
          setState({ data: null, loading: false, error: err.message, errorDetails: err.details || null, refreshing: false });
        });
    }

    function refreshData() {
      setState(function(prev) { return Object.assign({}, prev, { refreshing: true }); });
      fetch('/api/diversification-refresh', { method: 'POST' })
        .then(function(res) {
          if (!res.ok) throw new Error('Refresh failed');
          return res.json();
        })
        .then(function() {
          // Reload data after refresh
          loadData();
        })
        .catch(function(err) {
          setState(function(prev) { return Object.assign({}, prev, { refreshing: false, error: 'Refresh failed: ' + err.message }); });
        });
    }

    React.useEffect(function() {
      loadData();
    }, []);

    if (state.loading) {
      return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6' },
        e('div', { className: 'max-w-6xl mx-auto' },
          e('div', { className: 'card p-8 text-center' },
            e('div', { className: 'text-slate-400' }, 'Loading diversification data...')
          )
        )
      );
    }

    if (state.error) {
      var details = state.errorDetails || {};
      var errorType = details.errorType || 'UNKNOWN';
      var showRefreshButton = errorType === 'NO_SNAPSHOT' || errorType === 'STORAGE_ERROR';

      return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6' },
        e('div', { className: 'max-w-6xl mx-auto' },
          e('div', { className: 'card p-8 text-center' }, [
            e('h3', { key: 'h', className: 'text-lg font-semibold mb-3 ' + (errorType === 'NO_SNAPSHOT' ? 'text-amber-400' : 'text-rose-400') },
              errorType === 'NO_SNAPSHOT' ? 'No Data Snapshot Available' : 'Error Loading Data'),
            e('div', { key: 'm', className: 'text-slate-400 mb-4' }, state.error),
            showRefreshButton ? e('button', {
              key: 'refresh',
              className: 'bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-md font-medium transition ' + (state.refreshing ? 'opacity-50 cursor-not-allowed' : ''),
              onClick: refreshData,
              disabled: state.refreshing
            }, state.refreshing ? 'Refreshing from HubSpot...' : 'Fetch Data from HubSpot') : null,
            !showRefreshButton ? e('div', { key: 'note', className: 'text-xs text-slate-500 mt-4 p-3 bg-slate-800 rounded' },
              'Check Netlify Functions logs for details.') : null
          ])
        )
      );
    }

    var data = state.data;

    // Status indicator dot
    function StatusDot(status) {
      var colors = {
        green: 'bg-emerald-500',
        yellow: 'bg-amber-500',
        red: 'bg-rose-500'
      };
      return e('span', {
        className: 'inline-block w-3 h-3 rounded-full ml-2 ' + (colors[status] || 'bg-slate-500')
      });
    }

    // Funnel Tile component
    function FunnelTile(tile) {
      var isInfo = tile.isInformational;
      var cardClass = 'card p-5 ' + (isInfo ? 'border-dashed border-slate-600 bg-slate-800/30' : '');

      return e('div', { key: tile.id, className: cardClass }, [
        e('div', { key: 'header', className: 'flex items-center justify-between mb-2' }, [
          e('span', { key: 'label', className: 'text-sm font-medium text-slate-300' }, tile.label),
          !isInfo && tile.status ? StatusDot(tile.status) : null
        ]),
        e('div', { key: 'count', className: 'text-3xl font-bold text-white mb-2' }, tile.count),
        !isInfo ? e('div', { key: 'targets', className: 'text-xs text-slate-400 space-y-1' }, [
          e('div', { key: 'j20' }, 'July 20 target: ' + tile.count + ' / ' + tile.targetJuly20),
          e('div', { key: 'ye' }, 'Year-end target: ' + tile.targetYE)
        ]) : e('div', { key: 'subtitle', className: 'text-xs text-slate-500 italic' }, tile.subtitle)
      ]);
    }

    // Conversion Rate Card
    function ConversionCard(rate) {
      var pct = rate.percent.toFixed(1) + '%';
      var raw = rate.numerator + ' of ' + rate.denominator;

      return e('div', { key: rate.id, className: 'card p-5' }, [
        e('div', { key: 'label', className: 'text-sm font-medium text-slate-300 mb-2' }, rate.label),
        rate.subtitle ? e('div', { key: 'sub', className: 'text-xs text-slate-500 mb-2' }, rate.subtitle) : null,
        e('div', { key: 'pct', className: 'text-2xl font-bold text-white' }, pct),
        e('div', { key: 'raw', className: 'text-sm text-slate-400 mt-1' }, raw)
      ]);
    }

    // Format the last updated date nicely
    var lastUpdated = data.meta.lastUpdated ? new Date(data.meta.lastUpdated) : null;
    var lastUpdatedStr = lastUpdated ? lastUpdated.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : 'Unknown';

    return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6' },
      e('div', { className: 'max-w-6xl mx-auto space-y-6' }, [
        // Header
        e('div', { key: 'header', className: 'card p-6' }, [
          e('div', { key: 'title-row', className: 'flex items-start justify-between' }, [
            e('div', { key: 'titles' }, [
              e('h1', { key: 'title', className: 'text-xl font-bold text-white mb-1' }, 'Diversification Scorecard — Non-Broadcast Activity'),
              e('p', { key: 'subtitle', className: 'text-sm text-slate-400' }, 'Weekly snapshot from HubSpot. Targets per board mandate (July 20 / YE).')
            ]),
            e('button', {
              key: 'refresh-btn',
              className: 'bg-slate-700 hover:bg-slate-600 text-slate-200 px-4 py-2 rounded-md text-sm font-medium transition flex items-center gap-2 ' + (state.refreshing ? 'opacity-50 cursor-not-allowed' : ''),
              onClick: refreshData,
              disabled: state.refreshing
            }, state.refreshing ? 'Refreshing...' : 'Refresh Data')
          ]),
          data.pace ? e('div', { key: 'pace', className: 'text-xs text-slate-500 mt-3' },
            'Progress: ' + data.pace.percentComplete + '% through H1 (' + data.pace.daysElapsed + ' of ' + data.pace.daysToJuly20 + ' days to July 20)'
          ) : null
        ]),

        // Row 1: Funnel Tiles
        e('div', { key: 'tiles', className: 'grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-4' },
          data.tiles.map(function(tile) { return FunnelTile(tile); })
        ),

        // Row 2: Conversion Rates
        e('div', { key: 'rates-section' }, [
          e('h2', { key: 'h', className: 'text-lg font-semibold text-slate-200 mb-4' }, 'Conversion Rates'),
          e('div', { key: 'rates', className: 'grid grid-cols-1 md:grid-cols-3 gap-4' },
            data.conversionRates.map(function(rate) { return ConversionCard(rate); })
          )
        ]),

        // Footer with prominent snapshot date
        e('div', { key: 'footer', className: 'card p-5' }, [
          e('div', { key: 'updated', className: 'text-center mb-3' }, [
            e('div', { key: 'label', className: 'text-xs text-slate-500 uppercase tracking-wide' }, 'Snapshot Date'),
            e('div', { key: 'time', className: 'text-base font-medium text-slate-300 mt-1' }, lastUpdatedStr)
          ]),
          e('div', { key: 'links', className: 'text-center text-xs text-slate-500 pt-3 border-t border-slate-700' }, [
            data.meta.hubspotViewUrl ? e('a', {
              key: 'link',
              href: data.meta.hubspotViewUrl,
              target: '_blank',
              rel: 'noopener noreferrer',
              className: 'text-blue-400 hover:text-blue-300 underline'
            }, 'View deals in HubSpot') : null,
            e('span', { key: 'sep', className: 'mx-2' }, '·'),
            e('span', { key: 'schedule' }, 'Auto-refreshes every Monday at 6am ET')
          ])
        ])
      ])
    );
  }

  /* ------------------------ Standalone Active BoB Screen ------------------------ */
  function ActiveBoBScreen(){
    var _data = React.useState({ activeBob: [], loading: true, err: null });
    var data = _data[0], setData = _data[1];

    var _filters = React.useState({
      parentCompany: '', market: '', station: '', product: '', ae: '',
      psmFusion: '', psmSi: '', psmCi: '', paymentMethod: '', contractRenewalStatus: '',
      contractEndYears: []
    });
    var filters = _filters[0], setFilters = _filters[1];

    // Extract year from date string (handles MM/DD/YYYY, YYYY-MM-DD, etc.)
    function extractYear(dateStr) {
      if (!dateStr) return null;
      var match = dateStr.match(/\b(20\d{2})\b/);
      return match ? match[1] : null;
    }

    // Get unique years from contract end dates
    function getUniqueYears() {
      var years = {};
      data.activeBob.forEach(function(r) {
        var year = extractYear(r.contractEndDate);
        if (year) years[year] = true;
      });
      return Object.keys(years).sort();
    }

    var _sort = React.useState({ column: 'parentCompany', direction: 'asc' });
    var sort = _sort[0], setSort = _sort[1];

    React.useEffect(function(){
      var paths = [
        'https://docs.google.com/spreadsheets/d/e/2PACX-1vSx0Pp-5H60alDG7lTOneta10phn8QwLqXhnj0SSuAxobX5oaPj206IRFywm0BMBVPSOqCeKccE8KWY/pub?gid=414930090&single=true&output=csv',
        '/public/active_bob.csv','public/active_bob.csv','/active_bob.csv','active_bob.csv'
      ];
      loadCsvText(paths).then(function(txt){
        var parsed = Papa.parse(txt, { header: true, skipEmptyLines: true });
        var rows = (parsed.data || []).map(function(r){
          return {
            parentCompany: (r['Parent Company'] || r.parentCompany || '').trim(),
            market: (r['Market'] || r.market || '').trim(),
            station: (r['Station'] || r.station || '').trim(),
            product: (r['Product'] || r.product || '').trim(),
            ae: (r['AE'] || r.ae || '').trim(),
            psmFusion: (r['PSM Fusion'] || r.psmFusion || '').trim(),
            psmSi: (r['PSM SI'] || r.psmSi || '').trim(),
            psmCi: (r['PSM CI'] || r.psmCi || '').trim(),
            contractStartDate: (r['Contract Item Start Date'] || r.contractStartDate || '').trim(),
            contractEndDate: (r['Contract End Date'] || r.contractEndDate || '').trim(),
            paymentMethod: (r['Payment Method'] || r.paymentMethod || '').trim(),
            contractRenewalStatus: (r['Contract Renewal Status'] || r.contractRenewalStatus || '').trim(),
            annualContractValue: parseMoney(r['Annual Contract Value'] || r.annualContractValue || 0)
          };
        }).filter(function(r){ return r.parentCompany || r.station; });
        setData({ activeBob: rows, loading: false, err: null });
      }).catch(function(err){
        setData({ activeBob: [], loading: false, err: 'Failed to load data: ' + (err.message || err) });
      });
    }, []);

    var activeBob = data.activeBob;

    function getUniqueValues(field) {
      var vals = {};
      activeBob.forEach(function(r){ if(r[field]) vals[r[field]] = true; });
      return Object.keys(vals).sort();
    }

    function FilterSelect(props) {
      return e('select', {
        className: 'bg-slate-700 border border-slate-600 rounded px-2 py-1 text-xs text-slate-200 min-w-[120px]',
        value: filters[props.field] || '',
        onChange: function(ev) {
          var newFilters = Object.assign({}, filters);
          newFilters[props.field] = ev.target.value;
          setFilters(newFilters);
        }
      }, [
        e('option', { key: '', value: '' }, 'All ' + props.label),
        getUniqueValues(props.field).map(function(v) {
          return e('option', { key: v, value: v }, v);
        })
      ]);
    }

    var filteredData = activeBob.filter(function(r) {
      if (filters.parentCompany && r.parentCompany !== filters.parentCompany) return false;
      if (filters.market && r.market !== filters.market) return false;
      if (filters.station && r.station !== filters.station) return false;
      if (filters.product && r.product !== filters.product) return false;
      if (filters.ae && r.ae !== filters.ae) return false;
      if (filters.psmFusion && r.psmFusion !== filters.psmFusion) return false;
      if (filters.psmSi && r.psmSi !== filters.psmSi) return false;
      if (filters.psmCi && r.psmCi !== filters.psmCi) return false;
      if (filters.paymentMethod && r.paymentMethod !== filters.paymentMethod) return false;
      if (filters.contractRenewalStatus && r.contractRenewalStatus !== filters.contractRenewalStatus) return false;
      if (filters.contractEndYears.length > 0) {
        var rowYear = extractYear(r.contractEndDate);
        if (!rowYear || filters.contractEndYears.indexOf(rowYear) === -1) return false;
      }
      return true;
    });

    var sortedData = filteredData.slice().sort(function(a, b) {
      var aVal = a[sort.column] || '';
      var bVal = b[sort.column] || '';
      if (sort.column === 'annualContractValue') {
        aVal = a.annualContractValue || 0;
        bVal = b.annualContractValue || 0;
      }
      if (aVal < bVal) return sort.direction === 'asc' ? -1 : 1;
      if (aVal > bVal) return sort.direction === 'asc' ? 1 : -1;
      return 0;
    });

    var totalACV = filteredData.reduce(function(sum, r) { return sum + (r.annualContractValue || 0); }, 0);
    var activeContracts = filteredData.length;
    var uniqueParents = {};
    var uniqueStations = {};
    filteredData.forEach(function(r) {
      if (r.parentCompany) uniqueParents[r.parentCompany] = true;
      if (r.station) uniqueStations[r.station] = true;
    });

    function handleSort(column) {
      if (sort.column === column) {
        setSort({ column: column, direction: sort.direction === 'asc' ? 'desc' : 'asc' });
      } else {
        setSort({ column: column, direction: 'asc' });
      }
    }

    function SortIndicator(props) {
      if (sort.column !== props.column) return null;
      return e('span', { className: 'ml-1' }, sort.direction === 'asc' ? '▲' : '▼');
    }

    function SortableHeader(props) {
      return e('th', {
        key: props.column,
        className: 'px-2 py-2 text-left text-xs font-medium text-slate-400 cursor-pointer hover:text-slate-200 whitespace-nowrap',
        onClick: function() { handleSort(props.column); }
      }, [props.label, e(SortIndicator, { key: 's', column: props.column })]);
    }

    function clearFilters() {
      setFilters({
        parentCompany: '', market: '', station: '', product: '', ae: '',
        psmFusion: '', psmSi: '', psmCi: '', paymentMethod: '', contractRenewalStatus: '',
        contractEndYears: []
      });
    }

    // Multi-select year filter component
    function YearMultiSelect() {
      var years = getUniqueYears();
      var selectedYears = filters.contractEndYears;

      function toggleYear(year) {
        var newYears;
        if (selectedYears.indexOf(year) >= 0) {
          newYears = selectedYears.filter(function(y) { return y !== year; });
        } else {
          newYears = selectedYears.concat([year]);
        }
        setFilters(Object.assign({}, filters, { contractEndYears: newYears }));
      }

      return e('div', { className: 'flex flex-wrap items-center gap-1' }, [
        e('span', { key: 'label', className: 'text-xs text-slate-400 mr-1' }, 'End Year:'),
        years.map(function(year) {
          var isSelected = selectedYears.indexOf(year) >= 0;
          return e('button', {
            key: year,
            className: 'px-2 py-1 text-xs rounded ' +
              (isSelected
                ? 'bg-blue-600 text-white'
                : 'bg-slate-700 text-slate-300 hover:bg-slate-600'),
            onClick: function() { toggleYear(year); }
          }, year);
        })
      ]);
    }

    if (data.loading) {
      return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6' },
        e('div', { className: 'max-w-[2000px] mx-auto text-center py-20' },
          e('div', { className: 'text-lg text-slate-400' }, 'Loading Active Book of Business...')
        )
      );
    }

    if (data.err) {
      return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6' },
        e('div', { className: 'max-w-[2000px] mx-auto' },
          e('div', { className: 'card p-6 text-center text-red-400' }, data.err)
        )
      );
    }

    // Export to CSV function
    function exportToCsv() {
      var headers = ['Parent Company','Market','Station','Product','AE','PSM Fusion','PSM SI','PSM CI','Contract Start Date','Contract End Date','Payment Method','Renewal Status','Annual Contract Value'];
      var rows = sortedData.map(function(r) {
        return [
          r.parentCompany,
          r.market,
          r.station,
          r.product,
          r.ae,
          r.psmFusion,
          r.psmSi,
          r.psmCi,
          r.contractStartDate,
          r.contractEndDate,
          r.paymentMethod,
          r.contractRenewalStatus,
          r.annualContractValue
        ].map(function(val) {
          var str = String(val == null ? '' : val);
          if (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0) {
            return '"' + str.replace(/"/g, '""') + '"';
          }
          return str;
        }).join(',');
      });
      var csv = [headers.join(',')].concat(rows).join('\n');
      var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      var url = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.href = url;
      link.download = 'active_book_of_business_' + new Date().toISOString().slice(0,10) + '.csv';
      link.click();
      URL.revokeObjectURL(url);
    }

    if (!activeBob.length) {
      return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6' },
        e('div', { className: 'max-w-[2000px] mx-auto' },
          e('div', { className: 'card p-6 text-center text-slate-400' },
            'No Active Book of Business data available.'
          )
        )
      );
    }

    return e('main', { className: 'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800' },
      e('div', { className: 'max-w-[2000px] mx-auto px-4 lg:px-6 xl:px-8 py-6 space-y-6' }, [
        e('div', { key: 'header', className: 'card p-4 flex items-center justify-between' }, [
          e('h1', { key: 'title', className: 'text-xl font-semibold text-slate-100' }, 'Active Book of Business'),
          e('button', {
            key: 'export',
            className: 'bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded text-sm font-medium',
            onClick: exportToCsv
          }, 'Export CSV')
        ]),

        e('div', { key: 'kpis', className: 'grid grid-cols-2 lg:grid-cols-4 gap-4' }, [
          e('div', { key: 'acv', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Total ACV'),
            e('div', { className: 'text-2xl font-bold text-emerald-400' }, fmtCompact(totalACV))
          ]),
          e('div', { key: 'contracts', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Active Contracts'),
            e('div', { className: 'text-2xl font-bold text-blue-400' }, activeContracts.toLocaleString())
          ]),
          e('div', { key: 'parents', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Parent Companies'),
            e('div', { className: 'text-2xl font-bold text-purple-400' }, Object.keys(uniqueParents).length.toLocaleString())
          ]),
          e('div', { key: 'stations', className: 'card p-4' }, [
            e('div', { className: 'text-xs text-slate-400 mb-1' }, 'Stations'),
            e('div', { className: 'text-2xl font-bold text-orange-400' }, Object.keys(uniqueStations).length.toLocaleString())
          ])
        ]),

        e('div', { key: 'filters', className: 'card p-4' }, [
          e('div', { className: 'flex items-center justify-between mb-3' }, [
            e('h3', { key: 'h', className: 'text-sm font-semibold text-slate-300' }, 'Filters'),
            e('button', {
              key: 'clear',
              className: 'text-xs text-blue-400 hover:text-blue-300',
              onClick: clearFilters
            }, 'Clear All')
          ]),
          e('div', { className: 'flex flex-wrap gap-2 mb-3' }, [
            e(FilterSelect, { key: 'f1', field: 'parentCompany', label: 'Parent Company' }),
            e(FilterSelect, { key: 'f2', field: 'market', label: 'Market' }),
            e(FilterSelect, { key: 'f3', field: 'station', label: 'Station' }),
            e(FilterSelect, { key: 'f4', field: 'product', label: 'Product' }),
            e(FilterSelect, { key: 'f5', field: 'ae', label: 'AE' }),
            e(FilterSelect, { key: 'f6', field: 'psmFusion', label: 'PSM Fusion' }),
            e(FilterSelect, { key: 'f7', field: 'psmSi', label: 'PSM SI' }),
            e(FilterSelect, { key: 'f8', field: 'psmCi', label: 'PSM CI' }),
            e(FilterSelect, { key: 'f9', field: 'paymentMethod', label: 'Payment Method' }),
            e(FilterSelect, { key: 'f10', field: 'contractRenewalStatus', label: 'Renewal Status' })
          ]),
          e(YearMultiSelect, { key: 'yearFilter' })
        ]),

        e('div', { key: 'table', className: 'card p-4' }, [
          e('div', { className: 'flex items-center justify-between mb-3' }, [
            e('h3', { key: 'h', className: 'text-sm font-semibold text-slate-300' }, 'Active Contracts'),
            e('span', { key: 'count', className: 'text-xs text-slate-400' }, sortedData.length + ' records')
          ]),
          e('div', { className: 'overflow-x-auto' },
            e('table', { className: 'w-full text-sm' }, [
              e('thead', { key: 'thead' },
                e('tr', { className: 'border-b border-slate-600' }, [
                  e(SortableHeader, { key: 'h1', column: 'parentCompany', label: 'Parent Company' }),
                  e(SortableHeader, { key: 'h2', column: 'market', label: 'Market' }),
                  e(SortableHeader, { key: 'h3', column: 'station', label: 'Station' }),
                  e(SortableHeader, { key: 'h4', column: 'product', label: 'Product' }),
                  e(SortableHeader, { key: 'h5', column: 'ae', label: 'AE' }),
                  e(SortableHeader, { key: 'h6', column: 'psmFusion', label: 'PSM Fusion' }),
                  e(SortableHeader, { key: 'h7', column: 'psmSi', label: 'PSM SI' }),
                  e(SortableHeader, { key: 'h8', column: 'psmCi', label: 'PSM CI' }),
                  e(SortableHeader, { key: 'h9', column: 'contractStartDate', label: 'Start Date' }),
                  e(SortableHeader, { key: 'h10', column: 'contractEndDate', label: 'End Date' }),
                  e(SortableHeader, { key: 'h11', column: 'paymentMethod', label: 'Payment' }),
                  e(SortableHeader, { key: 'h12', column: 'contractRenewalStatus', label: 'Renewal Status' }),
                  e(SortableHeader, { key: 'h13', column: 'annualContractValue', label: 'ACV' })
                ])
              ),
              e('tbody', { key: 'tbody' },
                sortedData.slice(0, 500).map(function(r, i) {
                  return e('tr', { key: i, className: 'border-t border-slate-700/30 hover:bg-slate-700/20' }, [
                    e('td', { key: 'c1', className: 'px-2 py-2 whitespace-nowrap' }, r.parentCompany),
                    e('td', { key: 'c2', className: 'px-2 py-2 whitespace-nowrap' }, r.market),
                    e('td', { key: 'c3', className: 'px-2 py-2 whitespace-nowrap' }, r.station),
                    e('td', { key: 'c4', className: 'px-2 py-2 whitespace-nowrap' }, r.product),
                    e('td', { key: 'c5', className: 'px-2 py-2 whitespace-nowrap' }, r.ae),
                    e('td', { key: 'c6', className: 'px-2 py-2 whitespace-nowrap' }, r.psmFusion),
                    e('td', { key: 'c7', className: 'px-2 py-2 whitespace-nowrap' }, r.psmSi),
                    e('td', { key: 'c8', className: 'px-2 py-2 whitespace-nowrap' }, r.psmCi),
                    e('td', { key: 'c9', className: 'px-2 py-2 whitespace-nowrap' }, r.contractStartDate),
                    e('td', { key: 'c10', className: 'px-2 py-2 whitespace-nowrap' }, r.contractEndDate),
                    e('td', { key: 'c11', className: 'px-2 py-2 whitespace-nowrap' }, r.paymentMethod),
                    e('td', { key: 'c12', className: 'px-2 py-2 whitespace-nowrap' }, r.contractRenewalStatus),
                    e('td', { key: 'c13', className: 'px-2 py-2 whitespace-nowrap text-right font-medium text-emerald-400' }, fmtCurrency(r.annualContractValue))
                  ]);
                })
              )
            ])
          ),
          sortedData.length > 500 ? e('div', { className: 'text-xs text-slate-400 mt-2 text-center' }, 'Showing first 500 of ' + sortedData.length + ' records') : null
        ]),

        e('footer', { key: 'footer', className: 'text-xs text-slate-500 text-center py-4 border-t border-slate-700' },
          'Data sourced from Google Sheets • Auto-refreshes weekly'
        )
      ])
    );
  }

  /* ------------------------ Screens & Router ------------------------ */
  function HomeScreen(props){
    function Tile(title,desc,go,isOpen){
      return e('button',{className:'card p-6 text-left hover:-translate-y-0.5 transition',onClick:go},
        e('div',{className:'flex items-center gap-2'},
          e('div',{className:'text-lg font-semibold'},title),
          isOpen ? null : e('span',{className:'text-xs bg-slate-700 text-slate-300 px-2 py-0.5 rounded'},'Password')
        ),
        e('div',{className:'text-slate-400 text-sm mt-1'},desc)
      );
    }
    return e('main',{className:'min-h-screen bg-gradient-to-br from-slate-900 to-slate-800 p-6'},
      e('div',{className:'max-w-4xl mx-auto space-y-6'},
        e('h1',{className:'text-xl font-semibold text-slate-100'},'Futuri Dashboard — Choose a module'),
        e('div',{className:'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4'},[
          Tile('Diversification','Non-broadcast pipeline from HubSpot.',function(){props.onRoute('diversification');},false),
          Tile('Active Book of Business','Live contract data from Google Sheets.',function(){props.onRoute('active-bob');},true),
          Tile('Financials','Password protected financial data.',function(){props.onRoute('financials');},false),
          Tile('KPI Package','Data-driven KPIs from CSV.',function(){props.onRoute('kpi');},false),
          Tile('Premiere Pacing','Revenue pacing vs budget analysis.',function(){props.onRoute('pacing');},false)
        ])
      )
    );
  }

  function App(){
    var _r=React.useState('home'); var route=_r[0], setRoute=_r[1];
    function Nav(){
      return e('div',{className:'bg-slate-900 border-b border-slate-700 p-3 flex items-center justify-between'},
        e('button',{className:'text-sm text-slate-300 hover:text-white',onClick:function(){setRoute('home');}},'← Home'),
        e('div',{className:'text-sm text-slate-500'},'Futuri Dashboard')
      );
    }
    var body = (route==='diversification') ? e(PasswordGate,null,e(DiversificationScreen))
             : (route==='financials') ? e(PasswordGate,null,e(FinancialsScreen))
             : (route==='kpi')        ? e(PasswordGate,null,e(KPIScreen))
             : (route==='pacing')     ? e(PasswordGate,null,e(PacingScreen))
             : (route==='active-bob') ? e(ActiveBoBScreen)
             :                          e(HomeScreen,{onRoute:setRoute});
    return e('div',{className:'min-h-screen bg-slate-950'}, route!=='home'?e(Nav):null, body);
  }

  ReactDOM.render(e(App), document.getElementById('root'));
})();