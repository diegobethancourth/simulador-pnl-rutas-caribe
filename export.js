/* Excel export: builds a workbook with live formulas from the scenario on screen. */
(function () {
  const AZUL = 'FF0C447C', AZUL2 = 'FF1E5A9E', HUB = 'FFC77700', GRIS = 'FF7D838C';
  const P_AZUL = 'FFD7E5F4', P_AMBAR = 'FFFBE8C8', P_GRIS = 'FFE2E5E9', INPUT = 'FFFFF2CC', SUB = 'FFF2F4F7', LINE = 'FFB9C0C9';
  const MONEY = '"$"#,##0;[Red]-"$"#,##0';
  const MONEY2 = '"$"#,##0.00;[Red]-"$"#,##0.00';
  const PCT = '0.0%';
  const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
  const thin = { style: 'thin', color: { argb: LINE } };

  function loadLib() {
    if (window.ExcelJS) return Promise.resolve();
    return new Promise((ok, ko) => {
      const s = document.createElement('script');
      s.src = 'exceljs.min.js'; s.onload = ok; s.onerror = () => ko(new Error('No se pudo cargar la librería de Excel'));
      document.head.appendChild(s);
    });
  }

  async function build(st) {
    const p = st.p, on = st.on, val = st.val, R = st.R;
    const T = id => (on[id] ? val[id] : 0);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Simulador P&L de rutas'; wb.created = new Date(); wb.calcProperties.fullCalcOnLoad = true;

    /* ---------------- Supuestos ---------------- */
    const ws = wb.addWorksheet('Supuestos', { views: [{ state: 'frozen', ySplit: 3, showGridLines: false }], properties: { tabColor: { argb: 'FFF0A020' } } });
    ws.columns = [{ width: 52 }, { width: 16 }, { width: 58 }];
    ws.mergeCells('A1:C1');
    ws.getCell('A1').value = 'Supuestos del modelo · China → Barbados';
    ws.getCell('A1').font = { bold: true, size: 15, color: { argb: 'FFFFFFFF' } };
    ws.getCell('A1').fill = fill(AZUL); ws.getRow(1).height = 26;
    ws.mergeCells('A2:C2');
    ws.getCell('A2').value = `Escenario exportado: ${st.presetLabel} · ${new Date().toLocaleDateString('es-PA')} · Las celdas amarillas se pueden cambiar; el estado de resultados se recalcula solo.`;
    ws.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF5F6B78' } }; ws.getCell('A2').alignment = { wrapText: true, vertical: 'top' }; ws.getRow(2).height = 28;
    ws.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    ['Concepto', 'Valor', 'Nota'].forEach((h, i) => { const c = ws.getCell(3, i + 1); c.value = h; c.font = { bold: true, color: { argb: 'FF33414F' } }; c.fill = fill(P_GRIS); c.border = { bottom: thin }; });

    const NAMES = {};
    let r = 4;
    const sec = (t, color) => { r++; const c = ws.getCell(r, 1); c.value = t; c.font = { bold: true, size: 12, color: { argb: color || AZUL } }; ws.getRow(r).height = 20; r++; };
    const inp = (name, label, v, fmt, note) => {
      ws.getCell(r, 1).value = label;
      const c = ws.getCell(r, 2); c.value = v; c.numFmt = fmt; c.fill = fill(INPUT); c.border = { top: thin, bottom: thin, left: thin, right: thin }; c.font = { color: { argb: 'FF1D2733' } };
      if (note) { ws.getCell(r, 3).value = note; ws.getCell(r, 3).font = { size: 10, color: { argb: 'FF5F6B78' } }; }
      wb.definedNames.add(`Supuestos!$B$${r}`, name); NAMES[name] = r; r++;
    };
    const N0 = '#,##0', N2 = '#,##0.00', US = '"$"#,##0', D = '0" días"';
    sec('Producto y negocio');
    inp('Unidades', 'Unidades por contenedor 40HQ', p.units, N0);
    inp('FOB_unidad', 'Costo FOB por unidad (US$)', p.fob, '"$"#,##0.00');
    inp('Precio_unidad', 'Precio de venta por unidad en Barbados (US$)', p.price, '"$"#,##0.00');
    inp('Arancel', 'Arancel Barbados (CET CARICOM)', p.duty, PCT, 'Según la partida arancelaria del producto');
    inp('Levy', 'Levy ambiental', p.levy, PCT, 'Se cobra sobre el valor CIF');
    inp('Seguro', 'Seguro de carga (% sobre FOB + fletes)', p.ins, '0.00%');
    inp('Gastos_op', 'Gastos operativos (% de venta)', p.opex, PCT, 'Marketing, distribución local, overhead');
    inp('Margen_obj', 'Margen neto objetivo', p.target, PCT);
    inp('Costo_capital', 'Costo de capital anual', p.cap, PCT, 'Para el costo financiero del inventario en tránsito');
    inp('Cont_por_anio', 'Contenedores por año', st.perYear, N0, 'Para el margen neto anual');

    sec('Fletes y costos · ruta directa (naviera)', AZUL2);
    inp('Flete_directo', 'Tarifa naviera China → Bridgetown (US$/40HQ)', p.fDirect, US, 'Con transbordo de la naviera en Kingston, Cartagena o Caucedo');
    inp('Agente_directo', 'Honorario del agente', p.feeD, US);
    sec('Fletes y costos · hub Panamá (Zona Libre de Colón)', HUB);
    inp('Flete_hub_tramo1', 'Flete China → Colón, MIT (US$/40HQ)', p.fLeg1, US);
    inp('Feeder', 'Feeder Colón → Bridgetown (US$/40HQ)', p.fFeeder, US, 'POR COTIZAR');
    inp('ZLC_variable', 'Manejo ZLC por contenedor', p.zlcVar, US, 'Traslado puerto ↔ ZLC, almacén, DMC, fumigación · POR COTIZAR');
    inp('ZLC_fijo', 'Manejo ZLC fijo por embarque', p.zlcFix, US, 'Se reparte entre los contenedores del embarque');
    inp('Agente_hub_base', 'Forwarder: honorario por contenedor', 150, US);
    inp('Agente_hub_fijo', 'Forwarder: honorario fijo por embarque', p.feeHfix, US, 'Se reparte entre los contenedores del embarque');
    sec('Destino Bridgetown · ambas rutas');
    inp('Cargos_destino', 'THC, port charge, D/O y broker', p.dest, US, 'POR COTIZAR');
    inp('Dias_libres', 'Días libres en puerto', p.free, N0);
    inp('Demora_dia', 'Demora por día pasado de los libres, 40\' (US$)', p.dmr, US);
    sec('PNL original · referencia', GRIS);
    inp('Flete_original', 'Flete presupuestado en el PNL original', p.fOrig, US);

    sec('Tiempos de tránsito');
    inp('Transito_directo', 'Directo: días puerta a puerto', 52, N0, '~52 días según cotizaciones nov-2026');
    inp('Hub_dias_Colon', 'Hub: días hasta Colón', 31, N0);
    inp('Hub_dias_Colon_BGI', 'Hub: días de Colón a Bridgetown (ZLC + feeder)', 14, N0);
    inp('Canal_dias_extra', 'Días extra si hay restricción del Canal', 4, N0);
    inp('Peak_dias_extra', 'Días extra del feeder en temporada alta', 3, N0);
    inp('Hub_absorbe', 'Parte de la congestión que sufre el hub', 0.3, PCT, 'El forwarder reprograma; la naviera no');
    inp('PSS_factor', 'Parte de la tarifa directa que lleva el PSS de temporada alta', 0.35, PCT);
    inp('Merma_directo', 'Merma de la ruta directa vs. la del hub', 0.3, PCT, 'Menos manipuleo en la ruta directa');

    sec('Triggers del escenario (0 = apagado)', 'FFC0392B');
    inp('T_alza_flete', 'Alza de flete transpacífico', T('t1') / 100, PCT);
    inp('T_congestion', 'Congestión en puerto de transbordo (días)', T('t2'), N0);
    inp('T_temporada', 'Temporada alta del feeder (recargo)', T('t3') / 100, PCT);
    inp('T_canal', 'Restricción del Canal (recargo US$)', T('t4'), US);
    inp('T_retencion', 'Retención aduanal en Bridgetown (días)', T('t5'), N0);
    inp('T_reclasif', 'Reclasificación arancelaria (puntos de arancel)', T('t6') / 100, PCT);
    inp('T_merma', 'Merma o daño por manipuleo (hub)', T('t7') / 100, '0.0%');
    sec('Palancas', 'FF2E7D32');
    inp('L_contenedores', 'Contenedores por embarque en ZLC (1 = sin consolidar)', on.l1 ? val.l1 : 1, N0);
    inp('L_descuento', 'Descuento por volumen en China → Colón', on.l2 ? val.l2 / 100 : 0, PCT);

    /* ---------------- Estado de resultados ---------------- */
    const ps = wb.addWorksheet('Estado de resultados', { views: [{ state: 'frozen', ySplit: 4, xSplit: 1, showGridLines: false }], properties: { tabColor: { argb: AZUL } } });
    wb.views = [{ activeTab: 1 }];
    ps.columns = [{ width: 52 }, { width: 17 }, { width: 9 }, { width: 17 }, { width: 9 }, { width: 17 }, { width: 9 }, { width: 16 }];
    ps.mergeCells('A1:H1');
    ps.getCell('A1').value = 'Estado de resultados por contenedor 40HQ · China → Barbados';
    ps.getCell('A1').font = { bold: true, size: 15, color: { argb: 'FFFFFFFF' } }; ps.getCell('A1').fill = fill(AZUL); ps.getRow(1).height = 26;
    ps.mergeCells('A2:H2');
    ps.getCell('A2').value = `Escenario: ${st.presetLabel}. Todas las cifras salen de la hoja Supuestos (celdas amarillas).`;
    ps.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF5F6B78' } };
    const H = ['US$ por contenedor', 'Directo naviera', '% venta', 'Hub Panamá ZLC', '% venta', 'PNL original', '% venta', 'Hub − Directo'];
    const HC = ['FF33414F', AZUL2, 'FF5F6B78', HUB, 'FF5F6B78', GRIS, 'FF5F6B78', 'FF33414F'];
    H.forEach((h, i) => { const c = ps.getCell(4, i + 1); c.value = h; c.font = { bold: true, color: { argb: HC[i] } }; c.fill = fill(P_GRIS); c.alignment = { horizontal: i ? 'right' : 'left', wrapText: true }; c.border = { bottom: thin }; });
    ps.getRow(4).height = 30;

    const k = ['direct', 'hub', 'orig'], COLS = ['B', 'D', 'F'];
    const DR = { direct: '(Arancel+T_reclasif+Levy)', hub: '(Arancel+T_reclasif+Levy)', orig: '(Arancel+Levy)' };
    const RES = {}; // results from the page, for cached values
    k.forEach(key => {
      const x = R[key];
      const fr = Object.fromEntries(x.fr);
      RES[key] = {
        units: x.units, sales: x.sales, fob: -x.fobTot,
        f1: -(fr['Flete naviera China → Bridgetown'] || fr['Flete tramo 1 China → Colón'] || 0),
        zlc: -(fr['Manejo ZLC (traslado, almacén, DMC)'] || 0), feeder: -(fr['Feeder Colón → Bridgetown'] || 0),
        ins: -x.insurance, cif: -x.cif, duty: -x.duties, dest: -(x.dest + x.fee), arr: x.arr, hold: x.rel - x.arr, rel: x.rel,
        dem: -x.dem, carry: -x.carry, landed: -x.landed, gross: x.gross, opex: -x.opex, net: x.net,
        unitc: x.landed / x.units, freight: x.freight, fmax: x.fmax, slack: x.fmax - x.freight, year: x.net * st.perYear
      };
    });
    const F = {
      units: { direct: 'Unidades*(1-T_merma*Merma_directo)', hub: 'Unidades*(1-T_merma)', orig: 'Unidades' },
      sales: c => `${c}5*Precio_unidad`,
      fob: () => '-Unidades*FOB_unidad',
      f1: { direct: '-(Flete_directo*(1+T_alza_flete)+ROUND(Flete_directo*PSS_factor*T_temporada,0)+T_canal)', hub: '-(Flete_hub_tramo1*(1+T_alza_flete)*(1-L_descuento)+T_canal)', orig: '-Flete_original' },
      zlc: { direct: '0', hub: '-(ZLC_variable+ZLC_fijo/L_contenedores)', orig: '0' },
      feeder: { direct: '0', hub: '-Feeder*(1+T_temporada)', orig: '0' },
      ins: c => `Seguro*SUM(${c}7:${c}10)`,
      cif: c => `SUM(${c}7:${c}11)`,
      duty: (c, key) => `${c}12*${DR[key]}`,
      dest: { direct: '-(Cargos_destino+Agente_directo)', hub: '-(Cargos_destino+Agente_hub_base+Agente_hub_fijo/L_contenedores)', orig: '-(Cargos_destino+Agente_directo)' },
      arr: { direct: 'Transito_directo+IF(T_canal>0,Canal_dias_extra,0)+T_congestion', hub: 'Hub_dias_Colon+IF(T_canal>0,Canal_dias_extra,0)+ROUND(T_congestion*Hub_absorbe,0)+Hub_dias_Colon_BGI+IF(T_temporada>0,Peak_dias_extra,0)', orig: 'Transito_directo' },
      hold: { direct: 'T_retencion', hub: 'T_retencion', orig: '0' },
      rel: c => `${c}15+${c}16`,
      dem: c => `-MAX(0,${c}16-Dias_libres)*Demora_dia`,
      carry: c => `-Unidades*FOB_unidad*Costo_capital*${c}17/365`,
      landed: c => `${c}12+${c}13+${c}14+${c}18+${c}19`,
      gross: c => `${c}6+${c}20`,
      opex: c => `-${c}6*Gastos_op`,
      net: c => `${c}21+${c}22`,
      unitc: c => `-${c}20/${c}5`,
      freight: c => `-(${c}8+${c}9+${c}10)`,
      fmax: (c, key) => `(${c}6*(1-Gastos_op-Margen_obj)+${c}14+${c}18+${c}19)/((1+Seguro)*(1+${DR[key]}))-Unidades*FOB_unidad`,
      slack: c => `${c}26-${c}25`,
      year: c => `${c}23*Cont_por_anio`
    };
    const ROWS = [
      [5, 'units', 'Unidades vendibles (después de merma)', '#,##0', 'n'],
      [6, 'sales', 'Ventas', MONEY, 'sub'],
      [7, 'fob', 'Mercancía FOB', MONEY],
      [8, 'f1', 'Flete principal (naviera directa · China → Colón · presupuesto)', MONEY],
      [9, 'zlc', 'Manejo en Zona Libre de Colón', MONEY],
      [10, 'feeder', 'Feeder Colón → Bridgetown', MONEY],
      [11, 'ins', 'Seguro de carga', MONEY],
      [12, 'cif', 'Valor CIF (base de impuestos)', MONEY, 'sub'],
      [13, 'duty', 'Arancel + levy ambiental (sobre CIF)', MONEY],
      [14, 'dest', 'Cargos destino + agente', MONEY],
      [15, 'arr', 'Días hasta llegar a Bridgetown', '0', 'n'],
      [16, 'hold', 'Días de retención en puerto', '0', 'n'],
      [17, 'rel', 'Días hasta liberar la carga', '0', 'n'],
      [18, 'dem', 'Demoras en puerto', MONEY],
      [19, 'carry', 'Costo financiero del inventario en tránsito', MONEY],
      [20, 'landed', 'Costo aterrizado total', MONEY, 'sub'],
      [21, 'gross', 'Margen bruto', MONEY, 'tot'],
      [22, 'opex', 'Gastos operativos', MONEY],
      [23, 'net', 'Margen neto', MONEY, 'tot'],
      [24, 'unitc', 'Costo aterrizado por unidad', MONEY2, 'k'],
      [25, 'freight', 'Flete total puerta a puerta (dentro del CIF)', MONEY, 'k'],
      [26, 'fmax', 'Flete máximo para lograr el margen objetivo', MONEY, 'k'],
      [27, 'slack', 'Holgura (+) o exceso (−) frente al flete actual', MONEY, 'k'],
      [28, 'year', 'Margen neto anual (contenedores por año)', MONEY, 'k']
    ];
    for (const [row, key, label, fmt, kind] of ROWS) {
      const rw = ps.getRow(row);
      rw.getCell(1).value = label;
      k.forEach((rk, i) => {
        const c = COLS[i]; const f = F[key];
        const formula = typeof f === 'function' ? f(c, rk) : f[rk];
        const cell = rw.getCell(c); cell.value = { formula, result: RES[rk][key] }; cell.numFmt = fmt;
        if (!kind || kind === 'sub' || kind === 'tot') {
          const pc = rw.getCell(String.fromCharCode(c.charCodeAt(0) + 1));
          pc.value = { formula: `IFERROR(${c}${row}/${c}$6,"")`, result: RES[rk].sales ? RES[rk][key] / RES[rk].sales : 0 };
          pc.numFmt = PCT; pc.font = { size: 10, color: { argb: 'FF5F6B78' } };
        }
      });
      if (kind !== 'n' && !['unitc', 'freight', 'fmax'].includes(key)) { const d = rw.getCell('H'); d.value = { formula: `D${row}-B${row}`, result: RES.hub[key] - RES.direct[key] }; d.numFmt = fmt === MONEY2 ? MONEY2 : MONEY; }
      const bg = kind === 'sub' ? SUB : kind === 'tot' ? P_AZUL : kind === 'k' ? 'FFFAFBFC' : null;
      for (let col = 1; col <= 8; col++) {
        const cell = rw.getCell(col); cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDE1E6' } } };
        if (bg) cell.fill = fill(bg);
        if (kind === 'sub' || kind === 'tot') cell.font = Object.assign({}, cell.font, { bold: true });
        if (kind === 'n') cell.font = { size: 10, color: { argb: 'FF5F6B78' } };
      }
      if (kind === 'tot') rw.height = 20;
    }
    ps.getRow(24).getCell(1).border = { top: { style: 'thin', color: { argb: LINE } } };
    ps.addConditionalFormatting({ ref: 'B21:H23', rules: [
      { type: 'cellIs', operator: 'lessThan', formulae: ['0'], style: { font: { color: { argb: 'FFC0392B' }, bold: true } } },
      { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], style: { font: { color: { argb: 'FF2E7D32' }, bold: true } } }] });
    ps.getCell('A30').value = 'Lectura: cada US$1 de flete cuesta más de US$1 porque el arancel y el levy se cobran sobre el CIF.';
    ps.getCell('A30').font = { italic: true, size: 10, color: { argb: 'FF5F6B78' } };
    ps.getCell('A31').value = { formula: '"Multiplicador del flete en el P&L: "&TEXT((1+Seguro)*(1+Arancel+T_reclasif+Levy),"0.00")', result: 'Multiplicador del flete en el P&L: ' + ((1 + p.ins) * (1 + R.direct.dutyRate)).toFixed(2) };
    ps.getCell('A31').font = { bold: true, color: { argb: AZUL } };
    ps.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .4, right: .4, top: .5, bottom: .5, header: .2, footer: .2 } };

    /* ---------------- Sensibilidad ---------------- */
    const ss = wb.addWorksheet('Sensibilidad al flete', { views: [{ showGridLines: false }], properties: { tabColor: { argb: HUB } } });
    ss.columns = [{ width: 26 }, { width: 20 }, { width: 14 }, { width: 20 }, { width: 14 }];
    ss.mergeCells('A1:E1'); ss.getCell('A1').value = 'Margen neto según el flete total por contenedor';
    ss.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }; ss.getCell('A1').fill = fill(AZUL); ss.getRow(1).height = 24;
    ss.mergeCells('A2:E2'); ss.getCell('A2').value = 'Mismo producto y mismos cargos de destino del estado de resultados; solo cambia el flete total (dentro del CIF).';
    ss.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF5F6B78' } };
    ['Flete total (US$/40HQ)', 'Directo: margen neto', '% venta', 'Hub: margen neto', '% venta'].forEach((h, i) => { const c = ss.getCell(4, i + 1); c.value = h; c.font = { bold: true, color: { argb: i === 1 || i === 2 ? AZUL2 : i > 2 ? HUB : 'FF33414F' } }; c.fill = fill(P_GRIS); c.alignment = { horizontal: i ? 'right' : 'left' }; });
    const ER = "'Estado de resultados'!";
    const fl = []; for (let v = 4000; v <= 12000; v += 1000) fl.push(v);
    fl.forEach((v, i) => {
      const row = 5 + i;
      const a = ss.getCell(row, 1); a.value = v; a.numFmt = MONEY; a.fill = fill(INPUT);
      [['B', 'direct'], ['D', 'hub']].forEach(([col, rk]) => {
        const c = ER + col;
        const x = R[rk];
        const res = x.sales * (1 - p.opex) - ((x.fobTot + v) * (1 + p.ins) * (1 + x.dutyRate) + x.dest + x.fee + x.dem + x.carry);
        const cell = ss.getCell(`${col}${row}`);
        cell.value = { formula: `${c}$6*(1-Gastos_op)-((Unidades*FOB_unidad+$A${row})*(1+Seguro)*(1+${DR[rk]})-${c}$14-${c}$18-${c}$19)`, result: res };
        cell.numFmt = MONEY;
        const pcol = String.fromCharCode(col.charCodeAt(0) + 1);
        const pc = ss.getCell(`${pcol}${row}`); pc.value = { formula: `${col}${row}/${c}$6`, result: res / x.sales }; pc.numFmt = PCT;
      });
    });
    ss.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    ss.addConditionalFormatting({ ref: `B5:E${4 + fl.length}`, rules: [
      { type: 'cellIs', operator: 'lessThan', formulae: ['0'], style: { fill: fill('FFF6D5D1') } },
      { type: 'cellIs', operator: 'greaterThanOrEqual', formulae: ['0'], style: { fill: fill('FFD6EFDC') } }] });

    /* ---------------- Notas ---------------- */
    const ns = wb.addWorksheet('Notas y fuentes', { views: [{ showGridLines: false }] });
    ns.columns = [{ width: 120 }];
    const notes = [
      ['Cómo usar este archivo', 'h'],
      ['1. Cambie las celdas amarillas de la hoja Supuestos (producto, fletes cotizados, triggers y palancas).'],
      ['2. La hoja Estado de resultados se recalcula sola para las tres rutas: directo naviera, hub Panamá y PNL original.'],
      ['3. La hoja Sensibilidad al flete muestra hasta qué flete total cada ruta sigue dando margen neto positivo.'],
      [''],
      ['Cómo calcula el modelo', 'h'],
      ['• Valor CIF = mercancía FOB + todos los fletes y el manejo hasta Bridgetown + seguro. Barbados cobra arancel (CET CARICOM) y levy ambiental de 1% sobre el CIF.'],
      ['• El IVA de 17.5% no entra al P&L: es recuperable para el importador registrado.'],
      ['• Costo financiero = valor FOB × costo de capital × días hasta liberar la carga ÷ 365.'],
      ['• Demoras = días en Bridgetown por encima de los días libres × tarifa diaria del contenedor.'],
      ['• La consolidación reparte la parte fija del manejo en ZLC y del honorario del forwarder entre los contenedores del embarque.'],
      ['• El PNL original no lleva triggers ni palancas: es la referencia con el flete presupuestado.'],
      [''],
      ['Fuentes de las referencias (nov-2026)', 'h'],
      ['• Tarifas FCL Shekou/Yantian → Barbados, ~US$9,100–9,550 por 40HQ, ~52 días: 5688.cn'],
      ['• Shanghái → Manzanillo, Panamá, US$5,000 por 40HQ (salida 17-nov-2026): Goodhope Freight'],
      ['• Días libres y demora en Barbados: Maersk, información local Barbados'],
      ['• IVA 17.5% y arancel basado en el CET de CARICOM: PwC Tax Summaries'],
      ['• Feeder Colón → Bridgetown, manejo en ZLC y cargos de destino: supuestos POR COTIZAR.'],
      ['• Los datos del producto (unidades, FOB, precio) son de ejemplo hasta que se reemplacen.']
    ];
    notes.forEach(([t, h], i) => { const c = ns.getCell(i + 1, 1); c.value = t; c.alignment = { wrapText: true, vertical: 'top' }; c.font = h ? { bold: true, size: 13, color: { argb: AZUL } } : { size: 11 }; });

    addComparison(wb, st);
    wb.eachSheet(sh => sh.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, c => { c.font = Object.assign({ name: 'Calibri', size: 11 }, c.font || {}); })));
    return wb;
  }


  /* ---------------- Offers: template, upload, comparison sheet ---------------- */
  const OHDR = [['agente', 'Agente', 22], ['ruta', 'Ruta (Directo / Hub)', 14], ['transb', 'Transbordo / hub', 16], ['flete', 'Flete principal (US$/40HQ)', 14],
    ['recargos', 'Recargos BAF, PSS, peak (US$)', 13], ['feeder', 'Feeder Colón → Bridgetown (solo hub)', 14], ['zlc', 'Manejo ZLC (solo hub)', 12],
    ['destino', 'Cargos destino Bridgetown', 12], ['honor', 'Honorario del agente', 12], ['transito', 'Tránsito puerta a puerto (días)', 12],
    ['libres', 'Días libres en destino', 10], ['freq', 'Salidas por semana', 10], ['bl', 'Un solo B/L (Sí / No)', 11], ['validez', 'Validez de la tarifa', 14]];
  const COL = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const toCell = (o, k) => k === 'ruta' ? (o.ruta === 'hub' ? 'Hub' : 'Directo') : k === 'bl' ? (o.bl === 'no' ? 'No' : 'Sí') : (o[k] === null || o[k] === undefined || o[k] === '' ? null : o[k]);
  function inputHeader(sh, row, first) {
    OHDR.forEach(([, l], i) => { const c = sh.getCell(row, i + 1); c.value = l; c.font = { bold: true, color: { argb: 'FF33414F' } }; c.fill = fill(P_GRIS); c.alignment = { wrapText: true, vertical: 'bottom' }; c.border = { bottom: thin }; });
    sh.getRow(row).height = 44;
  }
  function validations(sh, r0, r1) {
    for (let r = r0; r <= r1; r++) {
      sh.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Directo,Hub"'] };
      sh.getCell(`M${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Sí,No"'] };
      for (let c = 4; c <= 12; c++) sh.getCell(r, c).numFmt = c >= 10 ? '0' : '"$"#,##0';
      for (let c = 1; c <= 14; c++) { const x = sh.getCell(r, c); x.fill = fill(INPUT); x.border = { bottom: { style: 'hair', color: { argb: 'FFDDE1E6' } }, right: { style: 'hair', color: { argb: 'FFDDE1E6' } } }; }
    }
  }
  async function buildTemplate() {
    const wb = new ExcelJS.Workbook(); wb.creator = 'Simulador P&L de rutas';
    const sh = wb.addWorksheet('Ofertas', { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }], properties: { tabColor: { argb: 'FFF0A020' } } });
    sh.columns = OHDR.map(([, , w]) => ({ width: w + 2 }));
    sh.mergeCells('A1:N1'); sh.getCell('A1').value = 'Plantilla de ofertas · flete China → Bridgetown, Barbados (contenedor 40HQ)';
    sh.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }; sh.getCell('A1').fill = fill(AZUL); sh.getRow(1).height = 24;
    sh.mergeCells('A2:N2'); sh.getCell('A2').value = 'Una fila por oferta. Montos en US$ por contenedor 40HQ. Deje en blanco lo que el agente no cotizó: el simulador lo completa con sus supuestos y lo marca. Ruta: "Directo" (B/L de la naviera con su transbordo) o "Hub" (Colón · Zona Libre · feeder).';
    sh.getCell('A2').alignment = { wrapText: true, vertical: 'top' }; sh.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF5F6B78' } }; sh.getRow(2).height = 30;
    inputHeader(sh, 4); validations(sh, 5, 34);
    const offers = (window.__offers && window.__offers.get()) || [];
    offers.filter(o => !o.ej).forEach((o, i) => OHDR.forEach(([k], j) => { sh.getCell(5 + i, j + 1).value = toCell(o, k); }));
    wb.eachSheet(x => x.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, c => { c.font = Object.assign({ name: 'Calibri', size: 11 }, c.font || {}); })));
    return wb;
  }
  async function parseUpload(file) {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
    const sh = wb.getWorksheet('Ofertas') || wb.getWorksheet('Comparativo de ofertas') || wb.worksheets[0];
    let hr = 0; sh.eachRow((row, n) => { if (!hr && String(row.getCell(1).value || '').trim().toLowerCase().startsWith('agente')) hr = n; });
    if (!hr) throw new Error('No encontré la fila de encabezados (columna A = "Agente").');
    const val = c => { let v = c.value; if (v && typeof v === 'object') v = 'result' in v ? v.result : ('richText' in v ? v.richText.map(t => t.text).join('') : v.text ?? null); return v === undefined ? null : v; };
    const out = [];
    for (let r = hr + 1; r <= sh.rowCount; r++) {
      const row = sh.getRow(r); const ag = val(row.getCell(1)); if (ag === null || String(ag).trim() === '') continue;
      const o = {};
      OHDR.forEach(([k], j) => {
        let v = val(row.getCell(j + 1));
        if (k === 'ruta') v = /hub|zlc|panam|col[oó]n/i.test(String(v || '')) ? 'hub' : 'direct';
        else if (k === 'bl') v = /^n/i.test(String(v || '')) ? 'no' : 'si';
        else if (['agente', 'transb', 'validez'].includes(k)) v = v === null ? '' : (v instanceof Date ? v.toLocaleDateString('es-PA') : String(v));
        else { v = (v === null || v === '') ? null : Number(String(v).replace(/[$,\s]/g, '')); if (v !== null && isNaN(v)) v = null; }
        o[k] = v;
      });
      out.push(o);
    }
    if (!out.length) throw new Error('La plantilla no tiene ofertas.');
    return out;
  }
  function addComparison(wb, st) {
    const offers = (window.__offers && window.__offers.get()) || []; if (!offers.length) return;
    const sh = wb.addWorksheet('Comparativo de ofertas', { views: [{ state: 'frozen', ySplit: 4, xSplit: 1, showGridLines: false }], properties: { tabColor: { argb: 'FF2E7D32' } } });
    const RES = [['O', 'Flete usado', 12], ['P', 'Feeder usado', 11], ['Q', 'Manejo ZLC usado', 11], ['R', 'Destino usado', 11], ['S', 'Honorario usado', 11], ['T', 'Tránsito usado (días)', 10], ['U', 'Días libres usados', 10],
      ['V', 'Fletes dentro del CIF', 13], ['W', 'Valor CIF', 13], ['X', 'Arancel + levy', 12], ['Y', 'Días hasta liberar', 10], ['Z', 'Demoras', 10], ['AA', 'Costo financiero', 11],
      ['AB', 'Costo aterrizado', 13], ['AC', 'Ventas', 12], ['AD', 'Margen bruto', 12], ['AE', 'Margen neto', 13], ['AF', '% margen neto', 10], ['AG', 'Costo por unidad', 11], ['AH', 'Flete máximo para objetivo', 13], ['AI', 'Holgura vs. flete máximo', 13], ['AJ', 'Ranking (margen neto)', 10]];
    sh.columns = OHDR.map(([, , w]) => ({ width: w + 1 })).concat(RES.map(x => ({ width: x[2] + 1 })));
    sh.mergeCells('A1:AJ1'); sh.getCell('A1').value = 'Comparativo de ofertas · mismo producto, mismos impuestos y mismo escenario';
    sh.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }; sh.getCell('A1').fill = fill(AZUL); sh.getRow(1).height = 24;
    sh.mergeCells('A2:AJ2'); sh.getCell('A2').value = `Escenario: ${st.presetLabel}. Columnas A–N: lo que cotizó cada agente (celdas amarillas, editables; en blanco = no cotizado). Columnas O–U: valor usado (cotizado o supuesto de la hoja Supuestos). Columnas V–AJ: resultado con fórmulas.`;
    sh.getCell('A2').alignment = { wrapText: true, vertical: 'top' }; sh.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF5F6B78' } }; sh.getRow(2).height = 30;
    [['A3:N3', 'Lo que cotizó el agente', P_AMBAR], ['O3:U3', 'Valor usado (cotizado o supuesto)', P_GRIS], ['V3:AJ3', 'Resultado por contenedor', P_AZUL]].forEach(([rg, t, c]) => { sh.mergeCells(rg); const x = sh.getCell(rg.split(':')[0]); x.value = t; x.font = { bold: true, color: { argb: AZUL } }; x.fill = fill(c); x.alignment = { horizontal: 'center' }; });
    inputHeader(sh, 4);
    RES.forEach(([c, l]) => { const x = sh.getCell(`${c}4`); x.value = l; x.font = { bold: true, color: { argb: 'FF33414F' } }; x.fill = fill(P_GRIS); x.alignment = { wrapText: true, vertical: 'bottom', horizontal: 'right' }; x.border = { bottom: thin }; });
    const n = offers.length, r0 = 5, r1 = 4 + n;
    validations(sh, r0, r1);
    const F = r => {
      const hub = `$B${r}="Hub"`;
      return {
        O: `IF(ISBLANK(D${r}),IF(${hub},Flete_hub_tramo1,Flete_directo),D${r})`,
        P: `IF(${hub},IF(ISBLANK(F${r}),Feeder,F${r}),0)`,
        Q: `IF(${hub},IF(ISBLANK(G${r}),ZLC_variable+ZLC_fijo/L_contenedores,G${r}),0)`,
        R: `IF(ISBLANK(H${r}),Cargos_destino,H${r})`,
        S: `IF(ISBLANK(I${r}),IF(${hub},Agente_hub_base+Agente_hub_fijo/L_contenedores,Agente_directo),I${r})`,
        T: `IF(ISBLANK(J${r}),IF(${hub},Hub_dias_Colon+Hub_dias_Colon_BGI,Transito_directo),J${r})`,
        U: `IF(ISBLANK(K${r}),Dias_libres,K${r})`,
        V: `IF(${hub},O${r}*(1+T_alza_flete)+T_canal+P${r}*(1+T_temporada)+Q${r},O${r}*(1+T_alza_flete)+ROUND(O${r}*PSS_factor*T_temporada,0)+T_canal)+N(E${r})`,
        W: `(Unidades*FOB_unidad+V${r})*(1+Seguro)`,
        X: `W${r}*(Arancel+T_reclasif+Levy)`,
        Y: `T${r}+IF(T_canal>0,Canal_dias_extra,0)+IF(${hub},ROUND(T_congestion*Hub_absorbe,0)+IF(T_temporada>0,Peak_dias_extra,0),T_congestion)+T_retencion`,
        Z: `MAX(0,T_retencion-U${r})*Demora_dia`,
        AA: `Unidades*FOB_unidad*Costo_capital*Y${r}/365`,
        AB: `W${r}+X${r}+R${r}+S${r}+Z${r}+AA${r}`,
        AC: `Unidades*(1-T_merma*IF(${hub},1,Merma_directo))*Precio_unidad`,
        AD: `AC${r}-AB${r}`,
        AE: `AD${r}-AC${r}*Gastos_op`,
        AF: `IFERROR(AE${r}/AC${r},0)`,
        AG: `AB${r}/(Unidades*(1-T_merma*IF(${hub},1,Merma_directo)))`,
        AH: `(AC${r}*(1-Gastos_op-Margen_obj)-R${r}-S${r}-Z${r}-AA${r})/((1+Seguro)*(1+Arancel+T_reclasif+Levy))-Unidades*FOB_unidad`,
        AI: `AH${r}-V${r}`,
        AJ: `RANK(AE${r},$AE$${r0}:$AE$${r1})`
      };
    };
    const sims = offers.map(o => window.__offers.sim(o, true));
    const ranks = sims.map(x => 1 + sims.filter(y => y.net > x.net).length);
    offers.forEach((o, i) => {
      const r = r0 + i, x = sims[i], p = st.p;
      OHDR.forEach(([k], j) => { sh.getCell(r, j + 1).value = toCell(o, k); });
      const hub = o.ruta === 'hub', d = k => (o[k] === null || o[k] === undefined || o[k] === '');
      const N = st.on.l1 ? st.val.l1 : 1;
      const used = { O: d('flete') ? (hub ? p.fLeg1 : p.fDirect) : +o.flete, P: hub ? (d('feeder') ? p.fFeeder : +o.feeder) : 0, Q: hub ? (d('zlc') ? p.zlcVar + p.zlcFix / N : +o.zlc) : 0,
        R: x.dest, S: x.fee, T: d('transito') ? (hub ? 45 : 52) : +o.transito, U: d('libres') ? p.free : +o.libres };
      const res = Object.assign(used, { V: x.freight, W: x.cif, X: x.duties, Y: x.rel, Z: x.dem, AA: x.carry, AB: x.landed, AC: x.sales, AD: x.gross, AE: x.net, AF: x.pctNet, AG: x.unit, AH: x.fmax, AI: x.slack, AJ: ranks[i] });
      const f = F(r);
      RES.forEach(([c]) => {
        const cell = sh.getCell(`${c}${r}`); cell.value = { formula: f[c], result: res[c] };
        cell.numFmt = ['T', 'U', 'Y', 'AJ'].includes(c) ? '0' : c === 'AF' ? PCT : c === 'AG' ? MONEY2 : MONEY;
        if (['O', 'P', 'Q', 'R', 'S', 'T', 'U'].includes(c)) cell.font = { color: { argb: 'FF5F6B78' } };
        if (c === 'AE' || c === 'AJ') cell.font = { bold: true };
      });
    });
    sh.addConditionalFormatting({ ref: `AE${r0}:AE${r1}`, rules: [
      { type: 'cellIs', operator: 'lessThan', formulae: ['0'], style: { fill: fill('FFF6D5D1'), font: { color: { argb: 'FFC0392B' }, bold: true } } },
      { type: 'cellIs', operator: 'greaterThanOrEqual', formulae: ['0'], style: { fill: fill('FFD6EFDC'), font: { color: { argb: 'FF2E7D32' }, bold: true } } }] });
    sh.addConditionalFormatting({ ref: `AJ${r0}:AJ${r1}`, rules: [{ type: 'cellIs', operator: 'equal', formulae: ['1'], style: { fill: fill('FFD6EFDC') } }] });
    // amber marking of blanks in quoted columns D–K where a default is used
    sh.addConditionalFormatting({ ref: `D${r0}:K${r1}`, rules: [{ type: 'expression', formulae: [`AND(ISBLANK(D${r0}),COLUMN(D${r0})<>5,OR(COLUMN(D${r0})<>6,$B${r0}="Hub"),OR(COLUMN(D${r0})<>7,$B${r0}="Hub"))`], style: { fill: fill('FFFBE8C8'), font: { color: { argb: 'FF7A4B00' } } } }] });
    sh.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  }
  async function downloadTemplate() {
    const say = t => window.__offers && window.__offers.msg(t);
    try { say('Preparando la plantilla…'); await loadLib(); const wb = await buildTemplate(); const buf = await wb.xlsx.writeBuffer();
      const r = await saveBlob('Plantilla_ofertas_China_Barbados.xlsx', new Blob([buf])); say(r === 'declined' ? 'Descarga cancelada' : 'Plantilla descargada');
    } catch (e) { say('No se pudo generar la plantilla: ' + (e.message || e)); }
  }
  async function uploadOffers(file) {
    const say = t => window.__offers && window.__offers.msg(t);
    try { say('Leyendo la plantilla…'); await loadLib(); const list = await parseUpload(file); window.__offers.set(list); say(`${list.length} ofertas cargadas`); }
    catch (e) { say('No se pudo leer el archivo: ' + (e.message || e)); }
  }
  window.__buildTemplate = async () => { await loadLib(); return (await buildTemplate()).xlsx.writeBuffer(); };
  window.__parseOffers = async (buf) => { await loadLib(); return parseUpload({ arrayBuffer: async () => buf }); };

  async function saveBlob(name, blob) {
    try {
      if (window.claude && typeof window.claude.use === 'function') {
        const d = await window.claude.use('downloads');
        if (d) { await d.save({ filename: name, data: blob }); return 'ok'; }
      }
    } catch (e) {
      if (e && e.code === 'declined') return 'declined';
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    return 'ok';
  }

  async function exportXlsx() {
    const msg = document.getElementById('xlsMsg');
    const say = t => { if (msg) msg.textContent = t; };
    try {
      say('Preparando el Excel…');
      await loadLib();
      const st = window.__sim.getState();
      const wb = await build(st);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const tag = ({ hoy: 'hoy', peak: 'temporada_alta', aduana: 'problema_aduana', cons: 'hub_consolidado' })[st.preset] || 'personalizado';
      const res = await saveBlob(`PNL_rutas_China_Barbados_${tag}.xlsx`, blob);
      say(res === 'declined' ? 'Descarga cancelada' : 'Excel descargado');
      setTimeout(() => say(''), 4000);
      return blob;
    } catch (e) { say('No se pudo generar el Excel: ' + (e.message || e)); }
  }
  window.__exportXlsx = exportXlsx;
  window.__buildXlsx = async () => { await loadLib(); const wb = await build(window.__sim.getState()); return wb.xlsx.writeBuffer(); };
  document.addEventListener('DOMContentLoaded', () => {
    const t = document.getElementById('ofTpl'); if (t) t.addEventListener('click', downloadTemplate);
    const f = document.getElementById('ofFile'); if (f) f.addEventListener('change', e => { const x = e.target.files[0]; if (x) uploadOffers(x); e.target.value = ''; });
    ['btnXls', 'btnXlsTop'].forEach(id => { const b = document.getElementById(id); if (b) b.addEventListener('click', () => { if (id === 'btnXlsTop' && window.__sim) window.__sim.showView('pnl'); exportXlsx(); }); });
  });
})();
