/** =========================================================
 * PAINEL GERAL — DANNA (limpo, rápido e padronizado)
 *  A:N  -> Produção por Colaborador x Produto x Mês + TOTAL ANO
 *           - Destaque: melhor mês (verde) e pior mês >0 (vermelho)
 *           - Mini sumário por colaborador (Total 2025 e Melhor mês) com cor destaque
 *  Q:S  -> Resumo por Colaborador x Produto (Total ANO) — sem repetir nome
 *  U:AI -> Metas Mensais 2025 (preenchimento manual) — sem repetir nome
 *  Menu -> Painel → Atualizar Painel Geral (Rápido)
 * ========================================================= */

const SHEET_BASE         = 'Base';
const SHEET_PAINEL_GERAL = 'Painel Geral';
const TIMEZONE_PG        = 'America/Sao_Paulo';
const ANO_ALVO_PG        = 2025;

// Inícios de blocos
const RESUMO_START_COL   = 17; // Q
const METAS_START_COL    = 21; // U
const METAS_TITLE_ROW    = 2;  // linha do título "Metas Mensais ..."

// Paleta Danna
const COLOR_DARK_BLUE  = '#0B2E6B'; // títulos
const COLOR_LIGHT_BLUE = '#EAF2FF'; // cabeçalhos
const COLOR_TITLE_BG   = COLOR_DARK_BLUE;
const COLOR_TITLE_FG   = '#FFFFFF';
const COLOR_HEAD_BG    = COLOR_LIGHT_BLUE;
const COLOR_HEAD_FG    = '#111111';
const COLOR_BODY_WHITE = '#FFFFFF';
const COLOR_BODY_ZEBRA = '#FAFCFF';
const COLOR_CARD       = '#F7F9FC';
const COLOR_TOTAL_BG   = '#E3E6EA';
const COLOR_DARK_GREEN = '#0B5A33'; // melhor mês (verde)
const COLOR_DARK_RED   = '#B00020'; // pior mês (vermelho)
const COLOR_SUMMARY_BG = '#D9ECFF'; // mini-sumário (destaque coerente com azul da marca)

// Produtos (nomes iguais ao cabeçalho da Base)
const PRODUTOS_VALIDOS = [
  'Swiss RE','Sob Medida','Pré Formatado','Vida Mensal','Vida Único',
  'Automóvel','Empresarial','Dental PF','Dental PJ','Saúde','Prev. Único','Prev. Mensal'
];

/* ========================= MENU (Refresh) ========================= */
function onOpen(){ instalarMenuRefresh(); }

function instalarMenuRefresh(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Refresh')
    .addItem('Atualizar Painel & Realizado', 'refreshPainelERealizado')
    .addToUi();
}

/** Executa atualização do Painel Geral e, em seguida, do Realizado */
function refreshPainelERealizado(){
  // 1) Painel Geral (rápido) — preserva METAS
  atualizarPainelGeralRapido();

  // 2) Realizado — panorama de metas (empresa e colaboradores)
  try {
    if (typeof construirRealizado === 'function') {
      construirRealizado();
    } else {
      SpreadsheetApp.getActive().toast('Função "construirRealizado" não encontrada. Cole o arquivo Realizado.gs.', 'Refresh', 5);
    }
  } catch (e){
    SpreadsheetApp.getActive().toast('Erro ao atualizar Realizado: ' + e, 'Refresh', 8);
    throw e;
  }

  SpreadsheetApp.getActive().toast('Refresh concluído ✅', 'Refresh', 3);
}


// Pode colar logo depois de atualizarPainelGeralRapido()
function atualizarTudo(){
  atualizarPainelGeralRapido();
  construirAnalitico();
  SpreadsheetApp.getActive().toast('Painel + Analítico atualizados ✅', 'Painel', 3);
}

/* ========================= ENTRADA PRINCIPAL ========================= */
function construirPainelGeral(){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName(SHEET_BASE);
  if (!base) throw new Error(`Aba "${SHEET_BASE}" não encontrada.`);

  // (re)cria a aba
  let sh = ss.getSheetByName(SHEET_PAINEL_GERAL);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_PAINEL_GERAL);

  // Título geral
  sh.getRange('A1').setValue('PAINEL GERAL')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(COLOR_TITLE_BG).setFontColor(COLOR_TITLE_FG);

  // Agrega base
  const {mapMensal, colaboradoresOrdenados} = PG_agregar_(base);

  // Produção
  PG_renderProducao_(sh, mapMensal, colaboradoresOrdenados);

  // Resumo
  PG_renderResumoColabProduto_(sh, mapMensal, RESUMO_START_COL, 3);

  // Metas
  PG_renderMetasMensais_(sh, colaboradoresOrdenados, METAS_TITLE_ROW, METAS_START_COL);

  // Visual comum
  PG_formatarEstiloLimpo_(sh);

  // carimbo
  const agora = Utilities.formatDate(new Date(), TIMEZONE_PG, 'dd/MM/yyyy HH:mm:ss');
  sh.getRange('A1').setNote('Atualizado em ' + agora + ' (TZ: ' + TIMEZONE_PG + ')');
}

/* ========================= ATUALIZAÇÃO RÁPIDA (menu) ========================= */
function atualizarPainelGeralRapido(){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName(SHEET_BASE);
  if (!base) throw new Error(`Aba "${SHEET_BASE}" não encontrada.`);
  let sh = ss.getSheetByName(SHEET_PAINEL_GERAL);
  if (!sh){ construirPainelGeral(); return; }

  const {mapMensal, colaboradoresOrdenados} = PG_agregar_(base);

  // Atualiza só o que muda no dia a dia: Produção e Resumo
  PG_renderProducao_(sh, mapMensal, colaboradoresOrdenados, /*fast*/ true);
  PG_renderResumoColabProduto_(sh, mapMensal, RESUMO_START_COL, 3, /*fast*/ true);

  // carimbo
  const agora = Utilities.formatDate(new Date(), TIMEZONE_PG, 'dd/MM/yyyy HH:mm:ss');
  sh.getRange('A1').setNote('Atualizado em ' + agora + ' (Rápido)');
  SpreadsheetApp.getActive().toast('Painel atualizado (rápido) ✅', 'Painel', 3);
}

/* ========================= BLOCO 1 — PRODUÇÃO (A:N) ========================= */
function PG_renderProducao_(sh, mapMensal, colaboradores, fast){
  let r = 3;
  const meses = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

  // Limpa só a área A3:N(last)
  const lastRow = sh.getLastRow();
  const clearRows = Math.max(0, lastRow - 2);
  if (clearRows){
    sh.getRange(3,1,clearRows,14)
      .clearContent().setBackground(null).setFontColor(null).setFontWeight('normal')
      .setBorder(false,false,false,false,false,false);
  }

  colaboradores.forEach(colab=>{
    // ===== TÍTULO do colaborador (A..N, mesclado) =====
    const title = sh.getRange(r, 1, 1, 14);
    title.mergeAcross();
    sh.getRange(r,1).setValue(colab);
    title.setBackground(COLOR_TITLE_BG).setFontColor(COLOR_TITLE_FG).setFontWeight('bold');
    r++;

    // ===== MINI SUMÁRIO (1 linha, com cor destaque e borda sutil) =====
    const sums = PG_sumMonthsForColab_(mapMensal[colab]||{});
    const totalAnoColab = sums.reduce((a,b)=>a+b,0);
    const bestIdx = sums.length ? sums.indexOf(Math.max.apply(null, sums)) : -1;
    const bestName = bestIdx>=0 ? meses[bestIdx] : '-';
    const mini = [
      'RESUMO', 'Total 2025 (R$):', totalAnoColab, '', 'Melhor mês:', bestName, (bestIdx>=0? sums[bestIdx] : '')
    ];
    while (mini.length < 14) mini.push('');
    const miniRg = sh.getRange(r,1,1,14).setValues([mini]);
    miniRg
      .setBackground(COLOR_SUMMARY_BG)
      .setFontWeight('bold');
    sh.getRange(r,3).setNumberFormat('"R$" #,##0.00');
    sh.getRange(r,7).setNumberFormat('"R$" #,##0.00');
    miniRg.setBorder(true,false,true,false,false,false,'#C9D6EE',SpreadsheetApp.BorderStyle.SOLID);
    r++;

    // ===== CABEÇALHO (borda inferior sutil) =====
    const header = ['Produto'].concat(meses).concat(['TOTAL ANO (R$)']);
    const headRg = sh.getRange(r,1,1,14)
      .setValues([header])
      .setBackground(COLOR_HEAD_BG).setFontColor(COLOR_HEAD_FG).setFontWeight('bold');
    headRg.setBorder(false,false,true,false,false,false,'#C9D6EE',SpreadsheetApp.BorderStyle.SOLID);
    r++;

    // ===== LINHAS (>0 no ano) =====
    const byProd = mapMensal[colab] || {};
    const prodsOrdenados = Object.keys(byProd)
      .map(p=>{
        const arr = byProd[p] || Array(12).fill(0);
        const tot = arr.reduce((a,b)=>a+(b||0),0);
        return {p, arr, tot};
      })
      .filter(x=>x.tot > 0)
      .sort((a,b)=> b.tot - a.tot);

    if (prodsOrdenados.length){
      const linhas = prodsOrdenados.map(x=>{
        const row = [x.p];
        for (let m=0;m<12;m++) row.push(x.arr[m] > 0 ? PG_round2_(x.arr[m]) : '');
        row.push(PG_round2_(x.tot));
        return row;
      });

      // Escreve linhas
      const body = sh.getRange(r,1,linhas.length,14).setValues(linhas);

      // Zebra + moeda
      const bgs = linhas.map((_, i)=> Array(14).fill(i % 2 ? COLOR_BODY_ZEBRA : COLOR_BODY_WHITE));
      body.setBackgrounds(bgs);
      sh.getRange(r,2,linhas.length,12).setNumberFormat('"R$" #,##0.00');
      sh.getRange(r,14,linhas.length,1).setNumberFormat('"R$" #,##0.00');

      // ===== Destaques por linha (produto): melhor mês (verde) e pior mês >0 (vermelho) =====
      for (let i=0; i<linhas.length; i++){
        const mesesVals = linhas[i].slice(1, 13).map(v => Number(v)||0); // 12 meses
        const maxVal = Math.max.apply(null, mesesVals);
        const positivos = mesesVals.filter(v=>v>0);
        const minVal = positivos.length ? Math.min.apply(null, positivos) : 0;

        const rowIdx = r + i;

        if (isFinite(maxVal) && maxVal>0){
          const maxIdx = mesesVals.indexOf(maxVal);    // 0..11
          const colMax = 2 + maxIdx;                   // B..M
          sh.getRange(rowIdx, colMax, 1, 1)
            .setBackground(COLOR_DARK_GREEN)
            .setFontColor('#FFFFFF')
            .setFontWeight('bold')
            .setNote(`Melhor mês do ano para ${linhas[i][0]}: R$ ${formatBR(maxVal)}`);
        }

        if (isFinite(minVal) && minVal>0){
          const minIdx = mesesVals.indexOf(minVal);    // 0..11
          const colMin = 2 + minIdx;                   // B..M
          sh.getRange(rowIdx, colMin, 1, 1)
            .setBackground(COLOR_DARK_RED)
            .setFontColor('#FFFFFF')
            .setFontWeight('bold')
            .setNote(`Pior mês do ano para ${linhas[i][0]}: R$ ${formatBR(minVal)}`);
        }
      }

      r += linhas.length;
    } else {
      sh.getRange(r,1,1,14).setValue('— Sem produção neste ano —').setFontStyle('italic');
      r++;
    }

    r++; // espaço entre colaboradores
  });
}

/* ========================= BLOCO 2 — RESUMO (Q:S) ========================= */
function PG_renderResumoColabProduto_(sh, mapMensal, startCol, startRow, fast){
  // limpa bloco (só conteúdo e estilos do bloco)
  const lastRow = sh.getLastRow();
  if (lastRow >= startRow){
    sh.getRange(startRow, startCol, lastRow-startRow+1, 3)
      .clearContent().setBackground(null).setFontColor(null).setFontWeight('normal');
  }

  // Cabeçalho (borda inferior sutil)
  const head = sh.getRange(startRow, startCol, 1, 3)
    .setValues([['Colaborador','Produto','Total 2025 (R$)']])
    .setBackground(COLOR_HEAD_BG).setFontColor(COLOR_HEAD_FG).setFontWeight('bold');
  head.setBorder(false,false,true,false,false,false,'#C9D6EE',SpreadsheetApp.BorderStyle.SOLID);

  let r = startRow + 1;
  const colaboradores = Object.keys(mapMensal).sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));

  colaboradores.forEach(c=>{
    // Título do colaborador
    sh.getRange(r, startCol, 1, 3)
      .setValues([[c.toUpperCase(), '', '']])
      .setBackground(COLOR_TITLE_BG).setFontColor(COLOR_TITLE_FG).setFontWeight('bold');
    r++;

    const byProd = mapMensal[c]||{};
    const arr = Object.keys(byProd).map(p=>{
      const tot = (byProd[p]||[]).reduce((a,b)=>a+(b||0),0);
      return {p, tot: PG_round2_(tot)};
    }).filter(x=>x.tot>0).sort((a,b)=>b.tot-a.tot);

    if (arr.length){
      const linhas = arr.map(x=>['', x.p, x.tot]);
      sh.getRange(r, startCol, linhas.length, 3).setValues(linhas).setBackground(COLOR_CARD);
      sh.getRange(r, startCol+1, linhas.length, 1).setFontWeight('bold');            // Produto
      sh.getRange(r, startCol+2, linhas.length, 1).setNumberFormat('"R$" #,##0.00'); // Total
      r += linhas.length;
    } else {
      sh.getRange(r, startCol, 1, 3).setValues([['', '— Sem produção —', '']]).setFontStyle('italic');
      r++;
    }

    r++; // espaço
  });
}

/* ========================= BLOCO 3 — METAS (U:AI) ========================= */
function PG_renderMetasMensais_(sh, colabs, startRow, startCol){
  // limpa bloco
  const lastRow = sh.getLastRow();
  if (lastRow >= startRow) {
    sh.getRange(startRow, startCol, Math.max(1,lastRow-startRow+1), 15)
      .clearContent().setBackground(null).setFontColor(null).setFontWeight('normal');
  }

  // Título
  sh.getRange(startRow, startCol).setValue('Metas Mensais 2025 (Preenchimento Manual)')
    .setBackground(COLOR_HEAD_BG).setFontWeight('bold').setFontColor(COLOR_HEAD_FG);

  const meses = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
  const header = ['Colaborador','Produto'].concat(meses).concat(['TOTAL']);
  const head = sh.getRange(startRow+1, startCol, 1, header.length)
    .setValues([header])
    .setBackground('#FFE5C0').setFontWeight('bold').setFontColor('#111');
  head.setBorder(false,false,true,false,false,false,'#C9D6EE',SpreadsheetApp.BorderStyle.SOLID);

  let r = startRow + 2;

  colabs.forEach(c=>{
    // título do colaborador (sem repetir nas linhas)
    sh.getRange(r, startCol, 1, header.length)
      .setValues([[c.toUpperCase(), ''].concat(Array(13).fill(''))])
      .setBackground(COLOR_TITLE_BG).setFontColor(COLOR_TITLE_FG).setFontWeight('bold');
    r++;

    // linhas de produtos (colaborador vazio)
    PRODUTOS_VALIDOS.forEach(p=>{
      const row = ['', p].concat(Array(12).fill('')).concat(['']);
      sh.getRange(r, startCol, 1, header.length).setValues([row]).setBackground(COLOR_CARD);
      r++;
    });

    // Total por colaborador
    const totalRow = ['TOTAL '+c,''].concat(
      meses.map((_,i)=>{
        const colLet = startCol+2+i;
        const firstR = r-PRODUTOS_VALIDOS.length;
        const lastR  = r-1;
        return `=SUM(${colToA1_(colLet)}${firstR}:${colToA1_(colLet)}${lastR})`;
      })
    ).concat([`=SUM(${colToA1_(startCol+2)}${r}:${colToA1_(startCol+13)}${r})`]);

    sh.getRange(r, startCol, 1, header.length).setValues([totalRow])
      .setBackground(COLOR_TOTAL_BG).setFontWeight('bold');

    sh.getRange(r, startCol+2, 1, 12).setNumberFormat('"R$" #,##0.00');
    sh.getRange(r, startCol+14,1, 1).setNumberFormat('"R$" #,##0.00');

    r++;
    sh.getRange(r, startCol).setValue(''); r++; // espaço
  });

  // formatos gerais das áreas de valores
  const lastRow2 = sh.getLastRow();
  const bodyRows = Math.max(0, lastRow2 - (startRow+1));
  if (bodyRows>0){
    sh.getRange(startRow+2, startCol+2, bodyRows, 12).setNumberFormat('"R$" #,##0.00');
    sh.getRange(startRow+2, startCol+14, bodyRows, 1).setNumberFormat('"R$" #,##0.00');
  }
}

/* ========================= AGREGAÇÃO DA BASE ========================= */
function PG_agregar_(baseSh){
  const lastRow = baseSh.getLastRow();
  const lastCol = baseSh.getLastColumn();
  const dados   = baseSh.getRange(1,1,lastRow,lastCol).getValues();

  if (!dados || dados.length < 2) throw new Error('Aba Base sem dados.');
  const hdr = dados[0].map(String);
  const H   = PG_headerMap_(hdr);

  const cStatus = H[PG_normKey_('Status')];
  const cAng    = H[PG_normKey_('Angariador')];
  const cData   = H[PG_normKey_('Data')];
  const cMes    = H[PG_normKey_('Mês')] ?? H[PG_normKey_('Mes')];

  if (cStatus==null || cAng==null || (cData==null && cMes==null))
    throw new Error('Base precisa de "Status", "Angariador" e "Data" (ou "Mês").');

  const prodColMap = PG_productColsFromHeader_(hdr);
  if (!Object.keys(prodColMap).length) throw new Error('Colunas de produtos não encontradas na Base.');

  const mapMensal = {};   // { colab: { prod: [12] } }
  const colabsSet = new Set();

  for (let i=1;i<dados.length;i++){
    const row = dados[i];
    const st  = String(row[cStatus]||'').toLowerCase();
    if (st.includes('cancel')) continue;

    let year, mIdx;
    if (cData!=null){
      const d = PG_toDate_(row[cData]); if (!d) continue;
      year = d.getFullYear(); mIdx = d.getMonth();
    } else {
      const m = Number(row[cMes]); if (!isFinite(m)||m<1||m>12) continue;
      year = ANO_ALVO_PG; mIdx = m-1;
    }
    if (year !== ANO_ALVO_PG) continue;

    const colab = String(row[cAng]||'').trim(); if (!colab) continue;
    colabsSet.add(colab);
    if (!mapMensal[colab]) mapMensal[colab] = {};

    for (const prod of PRODUTOS_VALIDOS){
      const col = prodColMap[prod]; if (col==null) continue;
      const v   = PG_toNumber_(row[col]); if (!v) continue;

      if (!mapMensal[colab][prod]) mapMensal[colab][prod] = Array(12).fill(0);
      mapMensal[colab][prod][mIdx] += PG_round2_(v);
    }
  }

  const colaboradoresOrdenados = [...colabsSet].sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
  return { mapMensal, colaboradoresOrdenados };
}

/* ========================= VISUAL PADRÃO DA ABA ========================= */
function PG_formatarEstiloLimpo_(sh){
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr<2 || lc<2) return;

  // Fonte/alinhamento
  sh.getRange(1,1,lr,lc)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(false);
  sh.getRange(1,1,lr,1).setHorizontalAlignment('left');
  if (lc>1) sh.getRange(1,2,lr,lc-1).setHorizontalAlignment('right');

  // Larguras
  const setW=(c,w)=>{ try{ sh.setColumnWidth(c,w); }catch(_){ } };
  setW(1,220);
  for (let c=2;c<=14;c++) setW(c,110);
  setW(17,220); setW(18,170); setW(19,140);
  setW(21,220); setW(22,180); for (let c=23;c<=34;c++) setW(c,110); setW(35,130);

  // Remove bordas gerais
  sh.getRange(1,1,lr,lc).setBorder(false,false,false,false,false,false);

  // Bordas sutis: cabeçalhos já recebem nas funções específicas
}

/* ========================= HELPERS ========================= */
function PG_headerMap_(hdr){ const m={}; hdr.forEach((h,i)=>m[PG_normKey_(h)]=i); return m; }
function PG_normKey_(v){ return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').replace(/[.()%]/g,'').trim().toLowerCase(); }
function PG_productColsFromHeader_(hdr){
  const H=PG_headerMap_(hdr), map={};
  PRODUTOS_VALIDOS.forEach(n=>{
    const k=PG_normKey_(n), alt=k.replace('prev. ','prev ');
    const idx=H[k] ?? H[alt];
    if (idx!=null) map[n]=idx;
  });
  return map;
}
function PG_toNumber_(v){
  if (typeof v==='number') return isFinite(v)?Number(v):0;
  const s=String(v||'').replace(/\s/g,'').replace(/[R$r$]/gi,'').replace(/\./g,'').replace(',', '.');
  const n=parseFloat(s); return isNaN(n)?0:n;
}
function PG_toDate_(val){
  if (val instanceof Date) return val;
  const s=String(val||'').trim(); if(!s) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) { const d=new Date(s); return isNaN(d)?null:d; }
  const m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m){ const d=new Date(+m[3],+m[2]-1,+m[1],m[4]?+m[4]:0,m[5]?+m[5]:0,m[6]?+m[6]:0); return isNaN(d)?null:d; }
  const d=new Date(s); return isNaN(d)?null:d;
}
function PG_round2_(n){ return Number((Number(n)||0).toFixed(2)); }
function colToA1_(c){ let s=''; let n=c; while(n>0){const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s; }
function PG_sumMonthsForColab_(byProd){
  const sums = Array(12).fill(0);
  Object.keys(byProd||{}).forEach(p=>{
    const arr = byProd[p]||[];
    for (let i=0;i<12;i++) sums[i] += Number(arr[i]||0);
  });
  return sums.map(PG_round2_);
}
function formatBR(n){
  return (Number(n)||0).toLocaleString('pt-BR',{minimumFractionDigits:2, maximumFractionDigits:2});
}
