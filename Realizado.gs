/** =========================================================
 * REALIZADO — Panorama de METAS (Empresa & Colaboradores)
 * - Aba: "Realizado"
 * - Seletores: Mês (B2), Ano (D2)
 * - Toggle: H2 "Mostrar por colaborador?" (checkbox em I2)
 * - Metas da Empresa: bloco "Metas Empresa 2025 (Manual)" no Painel Geral (AK:…)
 *   -> Se vazio, fallback = soma das metas dos colaboradores
 * - Campos: Meta Mensal, Meta Diária (planejada) e distribuição Semanal (por produto)
 * - Sem onEdit; atualização vem do menu "Refresh"
 * ========================================================= */

/* ===== Aliases globais (sem colidir com codigo.gs) ===== */
const _TZ_         = (typeof TIMEZONE_PG !== 'undefined' ? TIMEZONE_PG : 'America/Sao_Paulo');
const _PAINEL_     = (typeof SHEET_PAINEL_GERAL !== 'undefined' ? SHEET_PAINEL_GERAL : 'Painel Geral');
const _METAS_COL_  = (typeof METAS_START_COL !== 'undefined' ? METAS_START_COL : 21); // U
const _METAS_TIT_  = (typeof METAS_TITLE_ROW !== 'undefined' ? METAS_TITLE_ROW : 2);
const _METAS_HDR_  = _METAS_TIT_ + 1;
const _METAS_EMP_COL_ = (typeof METAS_EMP_START_COL !== 'undefined' ? METAS_EMP_START_COL : 37); // AK

const _PRODUTOS_ = (typeof PRODUTOS_VALIDOS !== 'undefined' && Array.isArray(PRODUTOS_VALIDOS) && PRODUTOS_VALIDOS.length)
  ? PRODUTOS_VALIDOS
  : ['Swiss RE','Sob Medida','Pré Formatado','Vida Mensal','Vida Único',
     'Automóvel','Empresarial','Dental PF','Dental PJ','Saúde','Prev. Único','Prev. Mensal'];

/* ===== Paleta herdada ===== */
const _COLOR_TITLE_BG   = (typeof COLOR_TITLE_BG   !== 'undefined') ? COLOR_TITLE_BG   : (typeof COLOR_DARK_BLUE  !== 'undefined' ? COLOR_DARK_BLUE  : '#0B2E6B');
const _COLOR_TITLE_FG   = (typeof COLOR_TITLE_FG   !== 'undefined') ? COLOR_TITLE_FG   : '#FFFFFF';
const _COLOR_HEAD_BG    = (typeof COLOR_HEAD_BG    !== 'undefined') ? COLOR_HEAD_BG    : (typeof COLOR_LIGHT_BLUE !== 'undefined' ? COLOR_LIGHT_BLUE : '#EAF2FF');
const _COLOR_HEAD_FG    = (typeof COLOR_HEAD_FG    !== 'undefined') ? COLOR_HEAD_FG    : '#111111';
const _COLOR_BODY_WHITE = (typeof COLOR_BODY_WHITE !== 'undefined') ? COLOR_BODY_WHITE : '#FFFFFF';
const _COLOR_BODY_ZEBRA = (typeof COLOR_BODY_ZEBRA !== 'undefined') ? COLOR_BODY_ZEBRA : '#FAFCFF';
const _COLOR_TOTAL_BG   = (typeof COLOR_TOTAL_BG   !== 'undefined') ? COLOR_TOTAL_BG   : '#E3E6EA';

const SHEET_REALIZADO = 'Realizado';

/* ======================= ENTRADA PRINCIPAL ======================= */
function construirRealizado(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_REALIZADO);
  if (!sh) sh = ss.insertSheet(SHEET_REALIZADO);

  const prevMesRaw = sh.getRange('B2').getValue();
  const prevAnoRaw = sh.getRange('D2').getValue();
  const prevToggle = sh.getRange('I2').getValue();

  RZ_layout_(sh, prevMesRaw, prevAnoRaw, prevToggle);

  const hoje = RZ_today_();
  const mesNome = String(sh.getRange('B2').getValue()||'').trim();
  const mes1 = RZ_monthNameToNumber_(mesNome) || (hoje.getMonth()+1);
  const ano  = Number(sh.getRange('D2').getValue()) || hoje.getFullYear();
  const showColab = Boolean(sh.getRange('I2').getValue());

  sh.getRange('H2').setValue('Mostrar por colaborador?').setFontWeight('bold');

  const duMes = RZ_businessDaysInMonth_(ano, mes1-1);
  sh.getRange('H3').setValue('Dias úteis no mês:').setFontWeight('bold');
  sh.getRange('I3').setValue(duMes);

  // ===== METAS: EMPRESA (preferência) =====
  const metasEmpresaMes = RZ_readMetasEmpresaMes_(mes1);       // {produto: valor}
  const metasEmpresaAno = RZ_readMetasEmpresaAno_();           // {produto: total}
  const usingEmpresaManual = Object.values(metasEmpresaMes).some(v=>RZ_toNumber_(v)>0) ||
                             Object.values(metasEmpresaAno).some(v=>RZ_toNumber_(v)>0);

  // ===== Fallback: somar colaboradores (se empresa manual não estiver preenchida) =====
  let metasColabMes = {}, colaboradoresOrdenados = [];
  if (!usingEmpresaManual){
    const res = RZ_coletarMetasPorColabProdutoMes_(mes1);
    metasColabMes = res.metasColabMes;
    colaboradoresOrdenados = res.colaboradoresOrdenados;
  } else {
    // Mesmo se usar empresa manual, ainda podemos exibir por colaborador se o toggle estiver ativo:
    const res = RZ_coletarMetasPorColabProdutoMes_(mes1);
    metasColabMes = res.metasColabMes;
    colaboradoresOrdenados = res.colaboradoresOrdenados;
  }

  // ===== Empresa efetiva: manual se disponível; senão soma colaboradores =====
  const metasEmpresaMesEfetiva = usingEmpresaManual ? metasEmpresaMes : RZ_agregarEmpresa_(metasColabMes);

  // Render: EMPRESA (mensal, diária, semanal)
  let r = 7;
  r = RZ_renderEmpresa_(sh, metasEmpresaMesEfetiva, duMes, mes1, ano, r, usingEmpresaManual);

  // (Opcional) Por colaborador
  if (showColab){
    r += 1;
    RZ_renderPorColaborador_(sh, metasColabMes, colaboradoresOrdenados, duMes, mes1, ano, r);
  }

  RZ_applyStyles_(sh);
  SpreadsheetApp.getActive().toast('Realizado (Panorama de Metas) atualizado ✅', 'Realizado', 3);
}

/* ======================= LAYOUT / TOGGLE ======================= */
function RZ_layout_(sh, prevMesRaw, prevAnoRaw, prevToggle){
  const lastRow = sh.getLastRow();
  if (lastRow > 6) sh.getRange(7,1,lastRow-6,80).clearContent().setBackground(null).setFontColor(null).setFontWeight('normal');
  sh.getRange('H2:I5').clearContent();

  sh.getRange('A1').setValue('REALIZADO — Panorama de Metas').setFontWeight('bold').setFontSize(12)
    .setBackground(_COLOR_TITLE_BG).setFontColor(_COLOR_TITLE_FG);

  sh.getRange('A2').setValue('Mês:').setFontWeight('bold');
  sh.getRange('C2').setValue('Ano:').setFontWeight('bold');

  const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(meses,true).setAllowInvalid(false).build());
  const y = (new Date()).getFullYear();
  const anos = [y-2,y-1,y,y+1,y+2].map(String);
  sh.getRange('D2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(anos,true).setAllowInvalid(false).build());

  if (!prevMesRaw) sh.getRange('B2').setValue(RZ_monthNumberToName_((new Date()).getMonth()+1));
  if (!prevAnoRaw) sh.getRange('D2').setValue(String((new Date()).getFullYear()));

  // Toggle "Mostrar por colaborador?"
  sh.getRange('H2').setValue('Mostrar por colaborador?').setFontWeight('bold');
  const i2 = sh.getRange('I2');
  if (i2.getDataValidation()==null) {
    i2.insertCheckboxes();
    i2.setValue(prevToggle === '' ? true : prevToggle); // default: ligado
  }

  sh.getRange('A3').setValue('Hoje:').setFontWeight('bold');
  sh.getRange('B3').setValue(RZ_fmtDate_(RZ_today_()));

  sh.setFrozenRows(6);
}

/* ======================= LEITURA DAS METAS DA EMPRESA (MANUAIS) ======================= */
// Estrutura esperada no Painel Geral a partir de AK (METAS_EMP_START_COL):
// Linha título (AK2): "Metas Empresa 2025 (Manual)"
// Cabeçalho (AK3): ['Produto','JAN','FEV',...'DEZ','TOTAL']
// Linhas (a partir de AK4): 1 por produto
function RZ_readMetasEmpresaMes_(mes1){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(_PAINEL_);
  if (!sh) return {};

  const mIdx = mes1 - 1; // 0..11
  const lastRow = sh.getLastRow();
  if (lastRow < _METAS_HDR_) return {};

  const width = 14; // Produto + 12 meses + TOTAL
  const values = sh.getRange(_METAS_HDR_, _METAS_EMP_COL_, lastRow - _METAS_HDR_ + 1, width).getValues();

  const metas = {};
  for (let i=1;i<values.length;i++){
    const row = values[i];
    const produto = String(row[0]||'').trim();
    if (!produto) continue;
    const vMes = RZ_toNumber_(row[1 + mIdx]); // meses começam na coluna 1 deste bloco
    if (!isNaN(vMes)) metas[produto] = vMes;
  }
  return metas;
}
function RZ_readMetasEmpresaAno_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(_PAINEL_);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  if (lastRow < _METAS_HDR_) return {};
  const width = 14;
  const values = sh.getRange(_METAS_HDR_, _METAS_EMP_COL_, lastRow - _METAS_HDR_ + 1, width).getValues();

  const metas = {};
  for (let i=1;i<values.length;i++){
    const row = values[i];
    const produto = String(row[0]||'').trim();
    if (!produto) continue;
    const total = RZ_toNumber_(row[13]); // TOTAL
    if (!isNaN(total)) metas[produto] = total;
  }
  return metas;
}

/* ======================= METAS POR COLABORADOR (Painel Geral) ======================= */
function RZ_coletarMetasPorColabProdutoMes_(mes1){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(_PAINEL_);
  if (!sh) return { metasColabMes:{}, colaboradoresOrdenados:[] };

  const mIdx = mes1 - 1; // 0..11
  const lastRow = sh.getLastRow();
  if (lastRow < _METAS_HDR_) return { metasColabMes:{}, colaboradoresOrdenados:[] };

  const width = 15; // Colab, Produto, 12 meses, TOTAL
  const values = sh.getRange(_METAS_HDR_, _METAS_COL_, lastRow - _METAS_HDR_ + 1, width).getValues();

  const metas = {};
  const colabsSet = new Set();
  let currentColab = null;

  for (let i=0;i<values.length;i++){
    const row = values[i];
    const c0 = String(row[0]||'').trim();
    const p1 = String(row[1]||'').trim();

    if (c0 && !p1){
      if (/^TOTAL\s/i.test(c0)) { currentColab = null; continue; }
      currentColab = c0;
      colabsSet.add(currentColab);
      if (!metas[currentColab]) metas[currentColab] = {};
      continue;
    }
    if (!p1 || !currentColab) continue;

    const vMes = RZ_toNumber_(row[2 + mIdx]) || 0;
    metas[currentColab][p1] = (metas[currentColab][p1]||0) + vMes;
  }

  const colaboradoresOrdenados = [...colabsSet].sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  return { metasColabMes: metas, colaboradoresOrdenados };
}

function RZ_agregarEmpresa_(metasColabMes){
  const empresa = {};
  _PRODUTOS_.forEach(p=>empresa[p]=0);
  Object.keys(metasColabMes||{}).forEach(c=>{
    const byProd = metasColabMes[c]||{};
    Object.keys(byProd).forEach(p=>{
      empresa[p] = (empresa[p]||0) + (RZ_toNumber_(byProd[p])||0);
    });
  });
  return empresa;
}

/* ======================= RENDER — EMPRESA ======================= */
function RZ_renderEmpresa_(sh, metasEmpresaMes, duMes, mes1, ano, startRow, usingEmpresaManual){
  let r = startRow;

  sh.getRange(r,1,1,8).mergeAcross()
    .setValue('EMPRESA — Metas (por produto)' + (usingEmpresaManual ? ' [META MANUAL]' : ' [SOMA COLAB.]'))
    .setBackground(_COLOR_TITLE_BG).setFontColor(_COLOR_TITLE_FG).setFontWeight('bold'); r++;

  // Mensal / Diária
  sh.getRange(r,1,1,5).setValues([[
    'Produto','Meta Mensal (R$)','Dias úteis (mês)','Meta Diária (R$)','Obs.'
  ]]).setBackground(_COLOR_HEAD_BG).setFontColor(_COLOR_HEAD_FG).setFontWeight('bold'); r++;

  const prods = Object.keys(metasEmpresaMes||{}).filter(p=>RZ_toNumber_(metasEmpresaMes[p])>0)
    .sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));

  const linhas = [];
  let tMensal=0, tDia=0;
  prods.forEach(p=>{
    const m = RZ_round2_(metasEmpresaMes[p]||0);
    const d = duMes>0 ? RZ_round2_(m/duMes) : 0;
    if (m>0){
      linhas.push([p, m, duMes, d, '']);
      tMensal += m; tDia += d;
    }
  });

  if (linhas.length){
    sh.getRange(r,1,linhas.length,5).setValues(linhas);
    sh.getRange(r,2,linhas.length,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(r,3,linhas.length,1).setNumberFormat('0');
    sh.getRange(r,4,linhas.length,1).setNumberFormat('"R$" #,##0.00');
    RZ_zebra_(sh, r, 1, linhas.length, 5);
    r += linhas.length;

    sh.getRange(r,1,1,5).setValues([['TOTAL', RZ_round2_(tMensal), duMes, RZ_round2_(tDia), '']])
      .setBackground(_COLOR_TOTAL_BG).setFontWeight('bold');
    sh.getRange(r,2,1,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(r,3,1,1).setNumberFormat('0');
    sh.getRange(r,4,1,1).setNumberFormat('"R$" #,##0.00');
    r += 2;
  } else {
    sh.getRange(r,1,1,5).setValue('— Sem metas definidas para este mês —').setFontStyle('italic'); r+=2;
  }

  // Semanas (planejado)
  const semanas = RZ_gerarSemanasUteis_(mes1, ano);
  sh.getRange(r,1,1,6).setValues([[
    'Semana','Início','Fim','Produto','Dias úteis','Meta da Semana (R$)'
  ]]).setBackground(_COLOR_HEAD_BG).setFontColor(_COLOR_HEAD_FG).setFontWeight('bold'); r++;

  const linhasW = [];
  const factorTotal = duMes>0 ? 1/duMes : 0;
  prods.forEach(p=>{
    const meta = RZ_toNumber_(metasEmpresaMes[p])||0;
    if (!meta) return;
    semanas.forEach((sem, idx)=>{
      const w = RZ_round2_(meta * (sem.bizDays * factorTotal));
      linhasW.push(['Semana '+(idx+1), sem.inicioBR, sem.fimBR, p, sem.bizDays, w]);
    });
  });

  if (linhasW.length){
    sh.getRange(r,1,linhasW.length,6).setValues(linhasW);
    sh.getRange(r,2,linhasW.length,2).setNumberFormat('dd/mm/yyyy');
    sh.getRange(r,5,linhasW.length,1).setNumberFormat('0');
    sh.getRange(r,6,linhasW.length,1).setNumberFormat('"R$" #,##0.00');
    RZ_zebra_(sh, r, 1, linhasW.length, 6);
    r += linhasW.length + 1;
  } else {
    sh.getRange(r,1,1,6).setValue('— Sem semanas úteis para distribuir —'); r+=2;
  }

  return r;
}

/* ======================= RENDER — POR COLABORADOR ======================= */
function RZ_renderPorColaborador_(sh, metasColabMes, colaboradores, duMes, mes1, ano, startRow){
  let r = startRow;

  sh.getRange(r,1,1,8).mergeAcross().setValue('POR COLABORADOR — Metas (por produto)')
    .setBackground(_COLOR_TITLE_BG).setFontColor(_COLOR_TITLE_FG).setFontWeight('bold'); r++;

  if (!colaboradores || !colaboradores.length){
    sh.getRange(r,1,1,6).setValue('— Nenhuma meta por colaborador para este mês —').setFontStyle('italic');
    return;
  }

  const semanas = RZ_gerarSemanasUteis_(mes1, ano);

  colaboradores.forEach(colab=>{
    const byProd = metasColabMes[colab] || {};
    const prods = Object.keys(byProd).filter(p=>RZ_toNumber_(byProd[p])>0)
      .sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));

    // Header do colaborador
    sh.getRange(r,1,1,8).mergeAcross()
      .setValue(colab)
      .setBackground(_COLOR_TITLE_BG).setFontColor(_COLOR_TITLE_FG).setFontWeight('bold');
    r++;

    // Mensal / Diária
    sh.getRange(r,1,1,5).setValues([[
      'Produto','Meta Mensal (R$)','Dias úteis (mês)','Meta Diária (R$)','Obs.'
    ]]).setBackground(_COLOR_HEAD_BG).setFontColor(_COLOR_HEAD_FG).setFontWeight('bold'); r++;

    const linhas = [];
    let tMensal=0, tDia=0;
    prods.forEach(p=>{
      const m = RZ_round2_(byProd[p]||0);
      const d = duMes>0 ? RZ_round2_(m/duMes) : 0;
      linhas.push([p, m, duMes, d, '']);
      tMensal += m; tDia += d;
    });

    if (linhas.length){
      sh.getRange(r,1,linhas.length,5).setValues(linhas);
      sh.getRange(r,2,linhas.length,1).setNumberFormat('"R$" #,##0.00');
      sh.getRange(r,3,linhas.length,1).setNumberFormat('0');
      sh.getRange(r,4,linhas.length,1).setNumberFormat('"R$" #,##0.00');
      RZ_zebra_(sh, r, 1, linhas.length, 5);
      r += linhas.length;

      sh.getRange(r,1,1,5).setValues([['TOTAL', RZ_round2_(tMensal), duMes, RZ_round2_(tDia), '']])
        .setBackground(_COLOR_TOTAL_BG).setFontWeight('bold');
      sh.getRange(r,2,1,1).setNumberFormat('"R$" #,##0.00');
      sh.getRange(r,3,1,1).setNumberFormat('0');
      sh.getRange(r,4,1,1).setNumberFormat('"R$" #,##0.00');
      r += 2;
    } else {
      sh.getRange(r,1,1,5).setValue('— Sem metas definidas para este colaborador no mês —').setFontStyle('italic'); r+=2;
    }

    // Semanas
    sh.getRange(r,1,1,6).setValues([[
      'Semana','Início','Fim','Produto','Dias úteis','Meta da Semana (R$)'
    ]]).setBackground(_COLOR_HEAD_BG).setFontColor(_COLOR_HEAD_FG).setFontWeight('bold'); r++;

    const linhasW = [];
    const factorTotal = duMes>0 ? 1/duMes : 0;
    prods.forEach(p=>{
      const meta = RZ_toNumber_(byProd[p])||0;
      semanas.forEach((sem, idx)=>{
        const w = RZ_round2_(meta * (sem.bizDays * factorTotal));
        linhasW.push(['Semana '+(idx+1), sem.inicioBR, sem.fimBR, p, sem.bizDays, w]);
      });
    });

    if (linhasW.length){
      sh.getRange(r,1,linhasW.length,6).setValues(linhasW);
      sh.getRange(r,2,linhasW.length,2).setNumberFormat('dd/mm/yyyy');
      sh.getRange(r,5,linhasW.length,1).setNumberFormat('0');
      sh.getRange(r,6,linhasW.length,1).setNumberFormat('"R$" #,##0.00');
      RZ_zebra_(sh, r, 1, linhasW.length, 6);
      r += linhasW.length + 2;
    } else {
      sh.getRange(r,1,1,6).setValue('— Sem semanas úteis para distribuir —'); r+=2;
    }
  });
}

/* ======================= SEMANAS / DIAS ÚTEIS ======================= */
function RZ_businessDaysInMonth_(year, month0){
  const d = new Date(year, month0, 1);
  let count = 0;
  while (d.getMonth() === month0){
    const day = d.getDay();
    if (day>=1 && day<=5) count++;
    d.setDate(d.getDate()+1);
  }
  return count;
}
function RZ_gerarSemanasUteis_(mes1, ano){
  const month0 = mes1-1;
  const first  = new Date(ano, month0, 1);
  const last   = new Date(ano, month0+1, 0);

  // Construir "semanas úteis" simples (seg–sex contíguos; feriados podem ser adicionados depois)
  const weeks = [];
  let cursor = new Date(first);

  // Semanas parciais antes da primeira segunda
  const pre = [];
  let d = new Date(first);
  while (d <= last && d.getDay() !== 1){ if (d.getDay()>=1 && d.getDay()<=5) pre.push(new Date(d)); d.setDate(d.getDate()+1); }
  if (pre.length) weeks.push(pre);

  // Semanas cheias e resto
  while (cursor <= last){
    if (cursor.getDay() === 1){
      const block = [];
      for (let k=0;k<5;k++){
        const dd = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate()+k);
        if (dd > last || dd.getMonth() !== month0) break;
        if (dd.getDay()>=1 && dd.getDay()<=5) block.push(dd);
      }
      if (block.length) weeks.push(block);
      cursor.setDate(cursor.getDate()+7);
    } else {
      cursor.setDate(cursor.getDate()+1);
    }
  }

  // Mapeia semana -> metadados
  return weeks.map(b=>{
    const ini=b[0], fim=b[b.length-1];
    return {
      inicioBR: Utilities.formatDate(ini, _TZ_, 'dd/MM/yyyy'),
      fimBR: Utilities.formatDate(fim, _TZ_, 'dd/MM/yyyy'),
      bizDays: b.length
    };
  });
}

/* ======================= HELPERS ======================= */
function RZ_today_(){
  const now = new Date();
  const s = Utilities.formatDate(now, _TZ_, 'yyyy-MM-dd HH:mm:ss');
  return new Date(s.replace(' ','T'));
}
function RZ_fmtDate_(d){ return Utilities.formatDate(d, _TZ_, 'dd/MM/yyyy'); }
function RZ_monthNameToNumber_(mes){
  const map={'JAN':1,'FEV':2,'MAR':3,'ABR':4,'MAI':5,'JUN':6,'JUL':7,'AGO':8,'SET':9,'OUT':10,'NOV':11,'DEZ':12};
  const k=String(mes||'').trim().toUpperCase().slice(0,3); return map[k]||null;
}
function RZ_monthNumberToName_(n){ const arr=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']; return (n>=1&&n<=12)?arr[n-1]:''; }
function RZ_toNumber_(v){
  if (typeof v==='number') return isFinite(v)?Number(v):0;
  const s=String(v||'').replace(/\s/g,'').replace(/[R$r$]/gi,'').replace(/\./g,'').replace(',', '.');
  const n=parseFloat(s); return isNaN(n)?0:n;
}
function RZ_round2_(n){ return Number((Number(n)||0).toFixed(2)); }
function RZ_applyStyles_(sh){
  const lr = sh.getLastRow(), lc = Math.max(8, sh.getLastColumn());
  if (lr<2) return;
  sh.getRange(1,1,lr,lc)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(false);
  sh.getRange(1,1,lr,1).setHorizontalAlignment('left');
  if (lc>1) sh.getRange(1,2,lr,lc-1).setHorizontalAlignment('right');

  const setW=(c,w)=>{ try{ sh.setColumnWidth(c,w); }catch(_){ } };
  setW(1,220); for (let c=2;c<=10;c++) setW(c,130);
}
function RZ_zebra_(sh, r, c, rows, cols){
  const bgs = [];
  for (let i=0;i<rows;i++){
    const color = (i%2===0)? _COLOR_BODY_WHITE : _COLOR_BODY_ZEBRA;
    bgs.push(Array(cols).fill(color));
  }
  sh.getRange(r,c,rows,cols).setBackgrounds(bgs);
}
