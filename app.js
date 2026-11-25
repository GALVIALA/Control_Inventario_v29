/* Control Inventario Offline: no network calls, data remains in browser storage. */
(() => {
  'use strict';

  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const Model = globalThis.InventoryModel;
  const Storage = globalThis.InventoryStorage;
  const Xlsx = globalThis.InventoryXlsx;
  const Lifecycle = globalThis.InventoryLifecycle;
  if (!Model || !Storage || !Xlsx || !Lifecycle) throw new Error('No se pudieron cargar los módulos locales del inventario.');
  const state = {
    workbook: null, records: [], filtered: [], activeLocations: {}, changes: {}, selectedSeries: {}, seriesIssues: [], consistencyReport: null,
    sessionId: null, inventory: null, readOnly: false, pendingCreate: null, pendingUpdate: null, returnContext: null, recoverySourceId: null, demo: false, commentKey: null, stream: null, scanTimer: null, scanControls: null, scanCandidate: null, scanHits: 0, scanCandidateAt: 0, ocrWorker: null, ocrWorkerPromise: null, ocrBusy: false, lastExport: null, lastExportMetadata: null, lastExportBytes: null, audioContext: null,
    exactCodes: new Map(), compactCodes: new Map(), viewingRecordKey: null, searchBaseline: '', searchRecordKey: null, suppressSearchGuard: false,
    missingLocationAcknowledged: new Set(), pendingNavigation: null, addRecordKey: null, pendingDeleteInventory: null, historyAdminMode: false, newInventoryFormOpen: false, persistPromise: Promise.resolve(), currentView: 'inventory', lastSavedAt: null, saveStatus: 'idle', saveError: null, sapLocationOptions: [], selectedSapLocation: ''
  };

  const ui = {
    importView: $('#importView'), historyView: $('#historyView'), inventoryView: $('#inventoryView'), locationReviewView: $('#locationReviewView'), statusView: $('#statusView'), finishView: $('#finishView'), bottomNav: $('#bottomNav'),
    inventoryName: $('#inventoryName'), demoExitButton: $('#demoExitButton'), activeInventoryCard: $('#activeInventoryCard'), activeInventoryName: $('#activeInventoryName'), activeInventoryProgress: $('#activeInventoryProgress'), homePrimaryActions: $('#homePrimaryActions'), startNewInventoryButton: $('#startNewInventoryButton'), newInventoryPanel: $('#newInventoryPanel'), newInventoryName: $('#newInventoryName'), fileInput: $('#fileInput'), updateFileInput: $('#updateFileInput'), updateFileLabel: $('#updateFileLabel'), restoreBackupButton: $('#restoreBackupButton'), restoreButton: $('#restoreButton'), historyButton: $('#historyButton'), manageHistory: $('#manageHistory'), historyList: $('#historyList'), historyEmpty: $('#historyEmpty'), historySearch: $('#historySearch'), demoButton: $('#demoButton'), recoveryLookupError: $('#recoveryLookupError'), recoveryLookupMessage: $('#recoveryLookupMessage'), retryRecoveryLookup: $('#retryRecoveryLookup'),
    searchInput: $('#searchInput'), results: $('#results'), empty: $('#emptyState'),
    countedTotal: $('#countedTotal'), saveState: $('#saveState'), scanButton: $('#scanButton'),
    scanDialog: $('#scanDialog'), cameraFrame: $('.camera-frame'), cameraVideo: $('#cameraVideo'), captureCanvas: $('#captureCanvas'), scannerHelp: $('#scannerHelp'), captureButton: $('#captureButton'),
    commentDialog: $('#commentDialog'), commentText: $('#commentText'), commentTitle: $('#commentTitle'),
    addCountDialog: $('#addCountDialog'), addCountMaterial: $('#addCountMaterial'), addCountCurrent: $('#addCountCurrent'), addCountValue: $('#addCountValue'), addCountTotal: $('#addCountTotal'),
    missingLocationDialog: $('#missingLocationDialog'), missingLocationMessage: $('#missingLocationMessage'),
    recoveryDialog: $('#recoveryDialog'), recoveryFileName: $('#recoveryFileName'), recoveryCompatible: $('#recoveryCompatible'), recoveryOmitted: $('#recoveryOmitted'), recoveryAmbiguous: $('#recoveryAmbiguous'),
    consistencyDialog: $('#consistencyDialog'), consistencyIntro: $('#consistencyIntro'), consistencyReport: $('#consistencyReport'), shareConsistencyButton: $('#shareConsistencyButton'),
    locationTrigger: $('#locationTrigger'), locationTriggerText: $('#locationTriggerText'), locationMenu: $('#locationMenu'), locationSearch: $('#locationSearch'), locationSuggestions: $('#locationSuggestions'), locationResults: $('#locationResults'), locationEmpty: $('#locationEmpty'), locationProgress: $('#locationProgress'), locationProgressLabel: $('#locationProgressLabel'), locationProgressCount: $('#locationProgressCount'), locationProgressPercent: $('#locationProgressPercent'), locationProgressBar: $('#locationProgressBar'),
    statusSapValue: $('#statusSapValue'), statusCountedValue: $('#statusCountedValue'), statusDifference: $('#statusDifference'), statusPercent: $('#statusPercent'), statusReviewed: $('#statusReviewed'), statusReviewedPercent: $('#statusReviewedPercent'), statusProgressBar: $('#statusProgressBar'), statusMatches: $('#statusMatches'), statusDifferences: $('#statusDifferences'), statusPending: $('#statusPending'), statusMissing: $('#statusMissing'), statusRelocated: $('#statusRelocated'), statusSeries: $('#statusSeries'), statusMissingToggle: $('#statusMissingToggle'), statusMissingPanel: $('#statusMissingPanel'), statusMissingPanelCount: $('#statusMissingPanelCount'), statusMissingRows: $('#statusMissingRows'), statusSeriesUnconfirmedToggle: $('#statusSeriesUnconfirmedToggle'), statusSeriesUnconfirmed: $('#statusSeriesUnconfirmed'), statusSeriesUnconfirmedPanel: $('#statusSeriesUnconfirmedPanel'), statusSeriesUnconfirmedPanelCount: $('#statusSeriesUnconfirmedPanelCount'), statusSeriesUnconfirmedRows: $('#statusSeriesUnconfirmedRows'), statusTopDifferences: $('#statusTopDifferences'), statusSavedAt: $('#statusSavedAt'),
    finishButton: $('#finishButton'), finishRows: $('#finishRows'), readonlyBanner: $('#readonlyBanner'),
    seriesIssuesBanner: $('#seriesIssuesBanner'), seriesIssuesBannerTitle: $('#seriesIssuesBannerTitle'), seriesIssuesBannerText: $('#seriesIssuesBannerText'),
    statusSeriesIssuesToggle: $('#statusSeriesIssuesToggle'), statusSeriesIssues: $('#statusSeriesIssues'), statusSeriesIssuesPanel: $('#statusSeriesIssuesPanel'), statusSeriesIssuesPanelCount: $('#statusSeriesIssuesPanelCount'), statusSeriesIssuesRows: $('#statusSeriesIssuesRows'),
    statusConsistencyToggle: $('#statusConsistencyToggle'), statusConsistencyCount: $('#statusConsistencyCount'), statusConsistencyPanel: $('#statusConsistencyPanel'), statusConsistencyPanelCount: $('#statusConsistencyPanelCount'), statusConsistencyContent: $('#statusConsistencyContent'),
    finishSeries: $('#finishSeries'), reviewCheck: $('#reviewCheck'), issuesReviewLabel: $('#issuesReviewLabel'), issuesReviewCheck: $('#issuesReviewCheck'), finishIssuesWarning: $('#finishIssuesWarning'), finishIssuesTitle: $('#finishIssuesTitle'), finishIssuesText: $('#finishIssuesText'), shareButton: $('#shareButton'),
    downloadButton: $('#downloadButton'), exportMessage: $('#exportMessage'), definitiveCloseSection: $('#definitiveCloseSection'), closeInventoryButton: $('#closeInventoryButton'), closeInventoryDialog: $('#closeInventoryDialog'), definitiveCloseCheck: $('#definitiveCloseCheck'), confirmCloseInventory: $('#confirmCloseInventory'), deleteInventoryDialog: $('#deleteInventoryDialog'), deleteInventoryName: $('#deleteInventoryName'), deleteInventoryCheck: $('#deleteInventoryCheck'), confirmDeleteInventory: $('#confirmDeleteInventory'), createInventoryDialog: $('#createInventoryDialog'), createInventorySummary: $('#createInventorySummary'), openInventoryWarning: $('#openInventoryWarning'), updatePreviewDialog: $('#updatePreviewDialog'), updatePreviewSummary: $('#updatePreviewSummary'), updatePreviewDetails: $('#updatePreviewDetails'), updateWarnings: $('#updateWarnings'), toast: $('#toast')
  };

  function normalize(v) {
    return String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[_\-\/]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function compactCode(v){return String(v??'').toUpperCase().replace(/[^A-Z0-9]/g,'');}
  function rebuildCodeIndex(){state.exactCodes=new Map();state.compactCodes=new Map();for(const record of state.records)for(const value of [record.material,record.barcode]){const code=String(value??'').trim();if(!code)continue;state.exactCodes.set(code.toUpperCase(),code);const compact=compactCode(code),list=state.compactCodes.get(compact)||[];if(!list.includes(code))list.push(code);state.compactCodes.set(compact,list);}}
  function resolveKnownCode(value,{ocr=false}={}){
    const raw=String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim(),clean=raw.replace(/[^A-Z0-9-]/g,'');
    for(const candidate of [raw,clean]){const exact=state.exactCodes.get(candidate);if(exact)return exact;}
    const compact=compactCode(raw),unique=state.compactCodes.get(compact);if(unique?.length===1)return unique[0];
    if(!ocr)return '';
    const variants=[];
    if(/^50\d{4,}$/.test(compact))variants.push(`S${compact.slice(1)}`);
    const swaps={O:'0',Q:'0',I:'1',L:'1',Z:'2',G:'6',B:'8',S:'5'};
    if(!/^S0/.test(compact))variants.push(compact.replace(/[OQILZGBS]/g,c=>swaps[c]));
    for(const candidate of variants){const exact=state.exactCodes.get(candidate);if(exact)return exact;const matches=state.compactCodes.get(compactCode(candidate));if(matches?.length===1)return matches[0];}
    return '';
  }
  function dictatedCodeValue(value){
    const digitWords={cero:'0',uno:'1',un:'1',una:'1',dos:'2',tres:'3',cuatro:'4',cinco:'5',seis:'6',siete:'7',ocho:'8',nueve:'9'};let spoken=String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
    spoken=spoken.replace(/\bguion\b/g,'-').replace(/\bdoble\s+(cero|uno|un|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve)\b/g,(_,word)=>digitWords[word]+digitWords[word]).replace(/\b(cero|uno|un|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve)\b/g,(_,word)=>digitWords[word]);
    const digits=(spoken.match(/\d/g)||[]).length;if(digits<4)return String(value??'');if(!/^\s*[a-z0-9][a-z0-9\s.,;:_/-]*\s*$/i.test(spoken))return String(value??'');const compact=spoken.replace(/[^a-z0-9-]/g,'').toUpperCase(),known=resolveKnownCode(compact);return known||compact;
  }
  function escapeHtml(v) { return String(v ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
  const escapeXml = Xlsx.escapeXml;
  const colFromRef = Xlsx.colFromRef;
  function colToNumber(col) { return [...col].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0); }
  function toast(message, duration = 2400) { const modal=$('dialog[open]');(modal||document.body).appendChild(ui.toast);ui.toast.textContent=message;ui.toast.classList.add('show');setTimeout(()=>{ui.toast.classList.remove('show');if(ui.toast.parentElement!==document.body)document.body.appendChild(ui.toast);},duration); }

  function saveStatusLabel(){
    if(state.saveStatus==='saving')return 'Guardando…';
    if(state.saveStatus==='error')return 'Error al guardar';
    if(state.lastSavedAt)return `Guardado · ${new Intl.DateTimeFormat('es-CL',{hour:'2-digit',minute:'2-digit'}).format(new Date(state.lastSavedAt))}`;
    return state.saveStatus==='saved'?'Guardado':'Guardado local';
  }
  function setSaveStatus(status,error=null){state.saveStatus=status;state.saveError=error;ui.saveState.textContent=status==='saving'?'Guardando…':status==='error'?'Error al guardar':status==='saved'?'Guardado':'Guardado local';ui.statusSavedAt.textContent=saveStatusLabel();}

  const db = Storage.createSessionStore(indexedDB);
  const saveQueue = Storage.createSerialSaveQueue(()=>persist(),(status,details)=>setSaveStatus(status,details.error));
  const inventorySessionId = workbook => Storage.inventorySessionId(workbook.records, crypto, Model, workbook.unmatchedSeries || []);
  const loadWorkbook = (bytes, fileName) => Xlsx.loadWorkbook(bytes, fileName, { JSZip, buildRecords });

  function buildRecords(wb) {
    const {inventory,seriesSheet}=wb, m=inventory.header.map, records=[], seriesItems=[], invalidPhysicalCountRows=[];
    if(seriesSheet){const sm=seriesSheet.header.map;for(const row of seriesSheet.rows.filter(r=>r.rowNumber>seriesSheet.header.rowNumber)){const material=String(row.values[sm.seriesMaterial]??'').trim(),serial=String(row.values[sm.serialNumber]??'').trim(),seriesLocation=String(row.values[sm.seriesLocation]??'').trim(),plant=String(row.values[sm.seriesPlant]??'').trim(),description=String(row.values[sm.seriesDescription]??'').trim();const originalStyle=row.cells[sm.serialNumber]?.style||0,styleState=Xlsx.seriesStyleState(wb.styleInfo,originalStyle);seriesItems.push({plant,material,description,serial,rowNumber:row.rowNumber,serialCol:sm.serialNumber,available:true,location:seriesLocation,originalStyle,styleState,selected:styleState==='selected'});}}
    for(const row of inventory.rows.filter(r=>r.rowNumber>inventory.header.rowNumber)){
      const material=String(row.values[m.material]??'').trim(); if(!material) continue;
      const stloc=String(row.values[m.stloc]??'').trim() || '—';
      const rawSerial=String(row.values[m.serialized]??''),serialControlRequired=/^(si|yes|y|x|1|true|serial)/i.test(normalize(rawSerial)),countState=Model.physicalCountState(row.values[m.count]);
      if(countState.status==='invalid')invalidPhysicalCountRows.push(row.rowNumber);
      records.push({rowNumber:row.rowNumber,plant:String(row.values[m.plant]??'').trim(),material,barcode:'',description:String(row.values[m.description]??'').trim(),batch:String(row.values[m.batch]??'').trim(),stloc,stockType:String(row.values[m.stockType]??'').trim(),
        sapQty:row.values[m.sapQty]??'',sapLocations:String(row.values[m.sapLocations]??'').trim(),unitValue:row.values[m.unitValue]??'',unitMeasure:String(row.values[m.unitMeasure]??'').trim(),
        count:countState.status==='valid'?countState.value:'',currentLocation:String(row.values[m.currentLocation]??'').trim(),comments:String(row.values[m.comments]??'').trim(),
        serialized:serialControlRequired,serialControlRequired,serials:[]});
    }
    wb.invalidPhysicalCountRows=invalidPhysicalCountRows;
    Model.assignStableRecordKeys(records);
    wb.seriesConsistencyReport=Model.seriesConsistencyReport(records,seriesItems);
    const ownership=Model.assignSeriesOwnership(records,seriesItems);
    for(const record of records)record.serialized=record.serialized||record.serials.length>0||record.alternateSerials.length>0;
    wb.unmatchedSeries=ownership.unresolved;
    wb.seriesIssues=ownership.issues;
    return records;
  }

  function recordSeriesItems(record){return [...(record.serials||[]),...(record.alternateSerials||[])];}
  function selectionsFromRecords(records){const selected={};for(const r of records){const list=recordSeriesItems(r).filter(s=>s.selected&&(!s.defaultRecordKey||s.defaultRecordKey===r.key)).map(Model.seriesSelectionKey);if(list.length)selected[r.key]=list;}return selected;}

  function formatConsistencyNumber(value){return new Intl.NumberFormat('es-CL',{maximumFractionDigits:2}).format(Number(value)||0);}
  function consistencyDifferenceText(difference){const amount=Math.abs(difference),noun=amount===1?'registro':'registros';return difference<0?`Faltan ${formatConsistencyNumber(amount)} ${noun} de serie`:`Sobran ${formatConsistencyNumber(amount)} ${noun} de serie`;}
  function consistencyReportMarkup(report,{copyButton=false}={}){
    const data=report||{seriesOnlyMaterials:[],quantityDifferences:[],locationDifferences:[],discrepancyCount:0};
    if(!data.discrepancyCount)return '<div class="consistency-pass"><strong>Planilla consistente</strong><span>No se detectaron diferencias entre Inventario y Series.</span></div>';
    const sections=[];
    if(data.seriesOnlyMaterials.length){
      const codes=data.seriesOnlyMaterials.map(item=>`<li><code>${escapeHtml(item.material)}</code><small>Plant ${escapeHtml(item.plants.join(', '))}</small></li>`).join('');
      sections.push(`<details class="consistency-group"><summary><div><strong>En Series, ausentes de Inventario</strong><span>${data.seriesOnlyMaterials.length} material${data.seriesOnlyMaterials.length===1?'':'es'}</span></div><i aria-hidden="true"></i></summary><div class="consistency-group-body">${copyButton?'<button class="copy-material-codes" type="button">Copiar códigos</button>':''}<ul class="consistency-code-list">${codes}</ul></div></details>`);
    }
    if(data.quantityDifferences.length){
      const rows=data.quantityDifferences.map(item=>`<li><div><strong>${escapeHtml(item.material)}</strong><small>Plant ${escapeHtml(item.plant)}</small></div><div><span>SAP ${escapeHtml(formatConsistencyNumber(item.sapQuantity))} · Series ${escapeHtml(formatConsistencyNumber(item.seriesCount))}</span><b>${escapeHtml(consistencyDifferenceText(item.difference))}</b></div></li>`).join('');
      sections.push(`<details class="consistency-group"><summary><div><strong>Diferencias de cantidad</strong><span>${data.quantityDifferences.length} material${data.quantityDifferences.length===1?'':'es'}</span></div><i aria-hidden="true"></i></summary><div class="consistency-group-body"><ul class="consistency-difference-list">${rows}</ul></div></details>`);
    }
    if(data.locationDifferences.length){
      const rows=data.locationDifferences.map(item=>`<li class="consistency-location-item"><div class="consistency-material-head"><strong>${escapeHtml(item.material)}</strong><small>Plant ${escapeHtml(item.plant)} · Total SAP ${escapeHtml(formatConsistencyNumber(item.sapQuantity))} · Series ${escapeHtml(formatConsistencyNumber(item.seriesCount))}</small></div><div class="consistency-location-table"><div><b>St Loc</b><b>SAP</b><b>Series</b></div>${item.locations.map(location=>`<div><span>${escapeHtml(location.location)}</span><span>${escapeHtml(formatConsistencyNumber(location.sapQuantity))}</span><span>${escapeHtml(formatConsistencyNumber(location.seriesCount))}</span></div>`).join('')}</div></li>`).join('');
      sections.push(`<details class="consistency-group"><summary><div><strong>Diferencias por St Loc</strong><span>${data.locationDifferences.length} material${data.locationDifferences.length===1?'':'es'}</span></div><i aria-hidden="true"></i></summary><div class="consistency-group-body"><ul class="consistency-location-list">${rows}</ul></div></details>`);
    }
    return sections.join('');
  }
  function consistencyReportText(report){
    const data=report||{seriesOnlyMaterials:[],quantityDifferences:[],locationDifferences:[]},lines=['Informe de consistencia Inventario / Series',''];
    lines.push('1. Materiales presentes en Series y ausentes de Inventario');
    if(data.seriesOnlyMaterials.length)for(const item of data.seriesOnlyMaterials)lines.push(`- ${item.material} · Plant ${item.plants.join(', ')}`);else lines.push('- Sin discrepancias');
    lines.push('','2. Diferencias entre cantidades SAP y registros de series');
    if(data.quantityDifferences.length)for(const item of data.quantityDifferences)lines.push(`- ${item.material} · Plant ${item.plant} · SAP ${formatConsistencyNumber(item.sapQuantity)} · Series ${formatConsistencyNumber(item.seriesCount)} · ${consistencyDifferenceText(item.difference)}`);else lines.push('- Sin discrepancias');
    lines.push('','3. Diferencias por St Loc');
    if(data.locationDifferences.length)for(const item of data.locationDifferences){lines.push(`- ${item.material} · Plant ${item.plant} · Total SAP ${formatConsistencyNumber(item.sapQuantity)} · Series ${formatConsistencyNumber(item.seriesCount)}`);for(const location of item.locations)lines.push(`  ${location.location}: SAP ${formatConsistencyNumber(location.sapQuantity)} · Series ${formatConsistencyNumber(location.seriesCount)} · ${consistencyDifferenceText(location.difference)}`);}else lines.push('- Sin discrepancias');
    lines.push('','Esta comprobación valida la coherencia de la planilla; no confirma las existencias físicas.');return lines.join('\n');
  }
  async function copyPlainText(text){try{await navigator.clipboard.writeText(text);}catch{const area=document.createElement('textarea');area.value=text;area.style.position='fixed';area.style.opacity='0';document.body.appendChild(area);area.select();document.execCommand('copy');area.remove();}}
  async function copyMaterialCodes(report){
    const codes=[...new Set((report?.seriesOnlyMaterials||[]).map(item=>item.material))].join('\n');if(!codes)return;
    await copyPlainText(codes);
    toast(`${codes.split('\n').length} código${codes.includes('\n')?'s':''} copiado${codes.includes('\n')?'s':''}`);
  }
  async function shareConsistencyReport(report){const text=consistencyReportText(report);if(navigator.share){try{await navigator.share({title:'Discrepancias del inventario',text});return;}catch(error){if(error?.name==='AbortError')return;}}await copyPlainText(text);toast('Informe completo copiado');}
  function renderConsistencyInto(container,report,{copyButton=false}={}){container.innerHTML=consistencyReportMarkup(report,{copyButton});$('.copy-material-codes',container)?.addEventListener('click',()=>copyMaterialCodes(report));}
  function reviewConsistency(report){
    ui.consistencyIntro.textContent=report.discrepancyCount?`Se detectaron ${report.discrepancyCount} discrepancias. Puedes corregir el Excel y volver a importarlo, o continuar con los materiales válidos.`:'La planilla superó las comprobaciones iniciales de consistencia.';
    $('#replaceConsistencyFile').textContent=report.discrepancyCount?'Corregir y volver a importar':'Volver a importar';renderConsistencyInto(ui.consistencyReport,report,{copyButton:true});ui.shareConsistencyButton.classList.toggle('hidden',!report.discrepancyCount);ui.shareConsistencyButton.onclick=()=>shareConsistencyReport(report);ui.consistencyDialog.returnValue='';ui.consistencyDialog.showModal();
    return new Promise(resolve=>ui.consistencyDialog.addEventListener('close',()=>resolve(ui.consistencyDialog.returnValue||'replace'),{once:true}));
  }

  const SERIES_ISSUE_TITLES={
    'serialized-without-series':'Material serializado sin series disponibles','series-without-material':'Serie sin material en Inventario','incompatible-location':'Serie disponible en otra ubicación','incompatible-plant':'Plant de serie incompatible','ambiguous-series':'Asociación de serie ambigua','incomplete-series-row':'Fila incompleta en Series','duplicate-series':'Serie duplicada en el mismo material','duplicate-series-def':'Serie repetida en material DEF','suspicious-series':'Número de serie sospechoso'
  };
  function criticalSeriesIssues(){return state.seriesIssues.filter(issue=>issue.severity==='critical');}
  function countingSeriesIssues(){const visibleWarnings=new Set(['serialized-without-series','duplicate-series-def','suspicious-series','incomplete-series-row']);return state.seriesIssues.filter(issue=>issue.severity==='critical'||visibleWarnings.has(issue.type));}
  function seriesIssueTitle(issue){return SERIES_ISSUE_TITLES[issue.type]||'Incidencia de serie';}
  function seriesIssueDetails(issue){
    const details=[];
    if(issue.material)details.push(`Material ${issue.material}`);
    if(issue.location)details.push(`ubicación ${issue.location}`);
    if(issue.inventoryLocations?.length)details.push(`St Loc Inventario ${issue.inventoryLocations.join(', ')}`);
    if(issue.serial)details.push(`serie ${issue.serial}`);
    if(issue.rowNumber)details.push(`fila Series ${issue.rowNumber}`);
    if(issue.inventoryRowNumber)details.push(`fila Inventario ${issue.inventoryRowNumber}`);
    if(issue.relatedRows?.length)details.push(`filas ${issue.relatedRows.join(', ')}`);
    if(issue.missingFields?.length)details.push(`faltan ${issue.missingFields.join(', ')}`);
    return details.join(' · ');
  }
  function recordIssueSummary(record){
    const visibleWarnings=new Set(['serialized-without-series','duplicate-series-def','suspicious-series','incomplete-series-row']);
    const issues=(record.seriesIssues||[]).filter(issue=>issue.severity==='critical'||visibleWarnings.has(issue.type));if(!issues.length)return '';
    const counts=new Map();for(const issue of issues)counts.set(issue.type,(counts.get(issue.type)||0)+1);
    const parts=[];
    for(const [type,count] of counts){const title=seriesIssueTitle({type});parts.push(count>1?`${count} × ${title.toLowerCase()}`:title);}
    return parts.slice(0,3).join(' · ')+(parts.length>3?` · ${parts.length-3} más`:'');
  }
  function renderSeriesIssues(){
    const visibleIssues=countingSeriesIssues(),count=visibleIssues.length,critical=visibleIssues.filter(issue=>issue.severity==='critical').length;
    ui.seriesIssuesBanner.classList.toggle('hidden',!count);
    if(count){ui.seriesIssuesBannerTitle.textContent=`${count} incidencia${count===1?'':'s'} de series`;ui.seriesIssuesBannerText.textContent=critical?`${critical} crítica${critical===1?'':'s'}. Los materiales válidos pueden seguir contándose.`:'Revisa los detalles; los materiales válidos pueden seguir contándose.';}
    ui.statusSeriesIssues.textContent=count;ui.statusSeriesIssuesPanelCount.textContent=count;ui.statusSeriesIssuesToggle.classList.toggle('hidden',!count);if(!count)ui.statusSeriesIssuesPanel.classList.add('hidden');
    ui.statusSeriesIssuesRows.innerHTML='';
    if(!count){ui.statusSeriesIssuesRows.innerHTML='<div class="status-list-empty">No hay incidencias de asociación de series.</div>';return;}
    for(const issue of visibleIssues){const row=document.createElement('div');row.className=`series-issue-row ${issue.severity==='critical'?'critical':''}`;row.innerHTML=`<div class="series-issue-row-head"><strong>${escapeHtml(seriesIssueTitle(issue))}</strong><b>${issue.severity==='critical'?'Crítica':'Advertencia'}</b></div><p>${escapeHtml(seriesIssueDetails(issue))}</p>`;ui.statusSeriesIssuesRows.appendChild(row);}
  }

  function renderConsistencyStatus(){const report=state.consistencyReport||{discrepancyCount:0};ui.statusConsistencyCount.textContent=report.discrepancyCount||0;ui.statusConsistencyPanelCount.textContent=report.discrepancyCount||0;renderConsistencyInto(ui.statusConsistencyContent,report,{copyButton:true});}

  function demoWorkbook() {
    const records=[
      {plant:'9999',material:'90000001',barcode:'90000001',description:'MOTOR DE DEMOSTRACIÓN',batch:'LOTE-DEMO-01',stloc:'D001',stockType:'(Un)Restricted-use',sapQty:3,sapLocations:'ESTANTE-A01',unitValue:'1.250,00',unitMeasure:'EA',serialized:false},
      {plant:'9999',material:'90000001',barcode:'90000001',description:'MOTOR DE DEMOSTRACIÓN',batch:'LOTE-DEMO-02',stloc:'D001',stockType:'Blocked',sapQty:1,sapLocations:'ESTANTE-B02',unitValue:'1.250,00',unitMeasure:'EA',serialized:false},
      {plant:'9999',material:'90000002',barcode:'90000002',description:'ACTUADOR DE DEMOSTRACIÓN',stloc:'D002',stockType:'(Un)Restricted-use',sapQty:2,sapLocations:'ESTANTE-C03',unitValue:'4.820,00',unitMeasure:'EA',serialized:true,serials:[{serial:'DEMO-SER-0001',available:true},{serial:'DEMO-SER-0002',available:true},{serial:'DEMO-SER-0003',available:false}]}
    ].map((r,i)=>({rowNumber:i+2,count:'',currentLocation:'',comments:'',serials:[],seriesIssues:[],...r}));
    Model.assignStableRecordKeys(records);
    return {fileName:'Demostración.xlsx',records,seriesIssues:[],seriesConsistencyReport:{seriesOnlyMaterials:[],quantityDifferences:[],locationDifferences:[],discrepancyCount:0},inventory:{header:{map:{material:'A',stloc:'C',count:'H',currentLocation:'I',comments:'J'}},path:'xl/worksheets/sheet1.xml'},seriesSheet:null};
  }

  function hasValidPhysicalCount(value){return Model.physicalCountState(value).status==='valid';}
  async function findRecoveryCandidates(currentWorkbook,sessionId){
    return Storage.evaluateSessionCandidates(await db.all(),currentWorkbook,sessionId,{
      loadWorkbook:saved=>loadWorkbook(saved.bytes,saved.fileName),
      inventorySessionId,
      recoveryPlan:Model.recoveryPlan,
      model:Model,
    });
  }
  function sessionEvaluationNotice(errors){if(!errors.length)return '';const first=errors[0],extra=errors.length>1?` y ${errors.length-1} más`:'';return `No se pudo evaluar la sesión “${first.fileName}”${extra}: ${first.message}`;}
  function confirmCompatibleRecovery(recovery){
    ui.recoveryFileName.textContent=recovery.saved.fileName||'conteo guardado';
    ui.recoveryCompatible.textContent=recovery.plan.compatibleCount;
    ui.recoveryOmitted.textContent=recovery.plan.omittedCount;
    ui.recoveryAmbiguous.textContent=recovery.plan.ambiguousCount;
    ui.recoveryDialog.returnValue='';
    ui.recoveryDialog.showModal();
    return new Promise(resolve=>ui.recoveryDialog.addEventListener('close',()=>resolve(ui.recoveryDialog.returnValue==='recover'),{once:true}));
  }
  function applyRecovery(currentRecords,recovery,{includeSavedWorkbook=false,mode='exact'}={}){
    if(!recovery)return {recoveredRows:0,skippedConflicts:0};
    const result=Model.recoveryData(currentRecords,recovery.workbook.records,recovery.saved,recovery.plan,{includeSavedWorkbook,mode});
    state.changes=result.changes;state.selectedSeries=result.selectedSeries;state.lastSavedAt=recovery.saved.updatedAt||null;
    return result;
  }

  async function prepareNewInventory(file) {
    const name=Lifecycle.assertInventoryName(ui.newInventoryName.value);
    if(!/\.xlsx$/i.test(file.name))throw new Error('Selecciona un archivo .xlsx.');
    const bytes=await file.arrayBuffer(),wb=await loadWorkbook(bytes,file.name),consistency=wb.seriesConsistencyReport||Model.seriesConsistencyReport(wb.records,[]);
    if(await reviewConsistency(consistency)!=='continue'){ui.fileInput.value='';toast('Corrige el Excel y vuelve a importarlo.');return;}
    const sourceFingerprint=await inventorySessionId(wb),activeSetting=await db.getSetting('activeInventoryId'),active=activeSetting?.value?await db.getInventory(activeSetting.value):null;
    state.pendingCreate={name,file,bytes,workbook:wb,sourceFingerprint};
    ui.createInventorySummary.textContent=`Se creará “${name}” con ${wb.records.length} materiales desde ${file.name}.`;
    ui.openInventoryWarning.classList.toggle('hidden',!active||active.status!=='open');
    ui.openInventoryWarning.innerHTML=active&&active.status==='open'?`<strong>Existe otro inventario abierto</strong><span>“${escapeHtml(active.name)}” permanecerá intacto. Debes confirmar el cambio de inventario activo.</span>`:'';
    ui.createInventoryDialog.showModal();
  }

  async function createPendingInventory(){
    const pending=state.pendingCreate;if(!pending)return;
    let inventory=Lifecycle.createInventory({name:pending.name,workbook:pending.workbook,sourceFingerprint:pending.sourceFingerprint,bytes:pending.bytes});
    inventory.selectedSeries=selectionsFromRecords(pending.workbook.records);
    await db.saveInventory(inventory,{activate:true});state.pendingCreate=null;ui.fileInput.value='';ui.newInventoryName.value='';
    await loadInventoryRecord(inventory,false);toast('Inventario creado y guardado');
  }

  async function loadInventoryRecord(inventory,readOnly=inventory.status==='closed'){
    if(!inventory?.bytes)throw new Error('El inventario no conserva una plantilla Excel recuperable.');
    const wb=await loadWorkbook(inventory.bytes,inventory.fileName),plan=Model.recoveryPlan(wb.records,wb.records),recovery={saved:inventory,workbook:wb,plan,progress:Storage.savedSessionProgress(wb,inventory,Model)};
    state.inventory=inventory;state.workbook=wb;state.records=wb.records;state.filtered=state.records;state.sessionId=inventory.id;state.recoverySourceId=null;state.demo=false;state.readOnly=readOnly;state.changes={};state.selectedSeries=selectionsFromRecords(wb.records);state.seriesIssues=wb.seriesIssues||[];state.consistencyReport=wb.seriesConsistencyReport||Model.seriesConsistencyReport(wb.records,[]);state.lastSavedAt=inventory.updatedAt||null;state.lastExport=null;state.lastExportMetadata=null;state.lastExportBytes=null;rebuildCodeIndex();applyRecovery(wb.records,recovery,{includeSavedWorkbook:true,mode:'exact'});enterInventory();
    ui.inventoryName.textContent=inventory.name;ui.readonlyBanner.classList.toggle('hidden',!readOnly);ui.definitiveCloseSection.classList.toggle('hidden',readOnly);ui.finishButton.disabled=readOnly;ui.searchInput.value=inventory.lastView?.search||'';render();
  }

  async function persist(bytes) {
    if(!state.inventory||state.demo||state.readOnly)return;
    const existing=await db.getInventory(state.inventory.id);if(!existing||existing.status!=='open')throw new Error('El inventario activo ya no admite cambios.');
    const savedAt=Date.now(),value={...existing,fileName:state.workbook.fileName,bytes:bytes||existing.bytes||state.workbook.bytes,changes:{...(existing.changes||{}),...state.changes},selectedSeries:{...(existing.selectedSeries||{}),...state.selectedSeries},updatedAt:savedAt,lastView:{view:state.currentView,search:ui.searchInput.value,scrollY:window.scrollY}};
    state.lastSavedAt=await Storage.confirmSave(()=>db.saveInventory(value),savedAt);state.inventory=value;return state.lastSavedAt;
  }
  async function persistNow(bytes){setSaveStatus('saving');try{const savedAt=await persist(bytes);setSaveStatus('saved');return savedAt;}catch(error){setSaveStatus('error',error);throw error;}}
  function queuePersist(){if(state.demo){setSaveStatus('saved');return Promise.resolve();}if(!state.sessionId)return Promise.resolve();state.persistPromise=saveQueue.enqueue();return state.persistPromise;}
  async function restoreLatest(){const setting=await db.getSetting('activeInventoryId'),inventory=setting?.value?await db.getInventory(setting.value):null;if(inventory)await loadInventoryRecord(inventory,false);}

  async function checkSavedSession(){
    ui.retryRecoveryLookup.disabled=true;
    try{await db.migrateLegacySessions();const setting=await db.getSetting('activeInventoryId'),inventory=setting?.value?await db.getInventory(setting.value):null,all=await db.allInventories();const active=inventory?.status==='open'?inventory:null,creating=state.newInventoryFormOpen;ui.restoreButton.disabled=!active;ui.homePrimaryActions.classList.toggle('hidden',creating);ui.newInventoryPanel.classList.toggle('hidden',!creating);ui.updateFileLabel.classList.toggle('hidden',!active||creating);ui.activeInventoryCard.classList.toggle('hidden',!active||creating);$('#homeTitle').textContent='Control de inventarios';$('#homeIntro').textContent=active&&!creating?'Los libros y conteos permanecen en este dispositivo.':creating?'Asigna un nombre y selecciona una copia de la plantilla Excel.':'Los libros y conteos permanecen en este dispositivo.';const backups=active?await db.backupsForInventory(active.id):[];ui.restoreBackupButton.classList.toggle('hidden',!backups.length||creating);if(active){ui.activeInventoryName.textContent=active.name;ui.activeInventoryProgress.textContent=`Actualizado ${new Intl.DateTimeFormat('es-CL',{dateStyle:'medium',timeStyle:'short'}).format(new Date(active.updatedAt))}`;ui.inventoryName.textContent=active.name;}else ui.inventoryName.textContent='Inventarios';ui.historyButton.disabled=!all.some(item=>item.id!==active?.id);ui.recoveryLookupError.classList.add('hidden');}
    catch(error){ui.restoreButton.classList.add('hidden');ui.recoveryLookupMessage.textContent='No fue posible consultar los conteos guardados en este dispositivo.';ui.recoveryLookupError.classList.remove('hidden');}
    finally{ui.retryRecoveryLookup.disabled=false;}
  }

  function hideMainViews(){for(const view of [ui.importView,ui.historyView,ui.inventoryView,ui.locationReviewView,ui.statusView,ui.finishView])view.classList.add('hidden');}
  function setActiveNav(id){for(const button of $$('.nav-button'))button.classList.toggle('active',button.id===id);}
  function enterInventory(){hideMainViews();ui.inventoryView.classList.remove('hidden');ui.bottomNav.classList.remove('hidden');ui.demoExitButton.classList.toggle('hidden',!state.demo);$('#manageInventoriesButton').classList.toggle('hidden',state.demo);state.currentView='inventory';setActiveNav('showInventory');populateLocationSelect();renderSeriesIssues();render();setTimeout(()=>prepareOcrWorker().catch(()=>{}),250);}
  function recordValue(record,field){return Object.prototype.hasOwnProperty.call(state.changes[record.key]||{},field)?state.changes[record.key][field]:record[field];}
  function setChange(record,field,value){if(state.readOnly){toast('El inventario histórico es de solo lectura');return;}state.changes[record.key]={...(state.changes[record.key]||{}),[field]:value};if(field==='currentLocation'||field==='count')state.missingLocationAcknowledged.delete(record.key);updateSummary();queuePersist();if(state.currentView==='locations')renderLocationReview();if(state.currentView==='status')renderStatus();}
  function groupRecords(records){const groups=new Map();for(const r of records){if(!groups.has(r.material))groups.set(r.material,[]);groups.get(r.material).push(r);}return [...groups.values()];}
  function formatValue(v){if(v===''||v==null)return '—';const n=Number(v);return Number.isFinite(n)&&String(v).trim()!==''?new Intl.NumberFormat('es-CL',{maximumFractionDigits:2}).format(n):String(v);}
  function stockTypeLabel(value){const raw=String(value??'').trim(),type=normalize(raw).replace(/[()]/g,'');if(!raw)return '';if(type.includes('blocked')||type.includes('bloqueado'))return 'Bloqueado';if(type.includes('unrestricted')||type.includes('restricted use')||type.includes('libre utilizacion'))return 'Libre';if(type.includes('quality')||type.includes('calidad'))return 'Calidad';return raw;}
  function batchSplitGroup(group){return group.some(record=>{const sameLocation=group.filter(item=>item.stloc===record.stloc),batches=new Set(sameLocation.map(item=>String(item.batch||'').trim()).filter(Boolean));return sameLocation.length>1&&batches.size>1;});}

  function activeRecordForGroup(group){const preferred=group.find(r=>r.stloc==='1140')||group[0],activeKey=state.activeLocations[group[0].material]||preferred.key;return group.find(r=>r.key===activeKey)||preferred;}
  function render(){
    const spoken=dictatedCodeValue(ui.searchInput.value),resolved=resolveKnownCode(spoken)||resolveKnownCode(ui.searchInput.value),rawQ=normalize(ui.searchInput.value),q=normalize(resolved||spoken),exact=resolved?state.records.filter(r=>[r.material,r.barcode].some(v=>String(v??'').toUpperCase()===resolved.toUpperCase())):q?state.records.filter(r=>[r.material,r.barcode].some(v=>{const n=normalize(v);return n===q||n===rawQ;})):[];
    state.filtered=q?(exact.length?exact:state.records.filter(r=>[r.material,r.barcode,r.description,r.stloc].some(v=>{const n=normalize(v);return n.includes(q)||(rawQ!==q&&n.includes(rawQ));}))):[];
    const groups=groupRecords(state.filtered).slice(0,40);state.viewingRecordKey=exact.length&&groups.length===1?activeRecordForGroup(groups[0]).key:null;
    ui.empty.classList.toggle('hidden',state.filtered.length>0);ui.results.innerHTML='';
    const emptyTitle=$('strong',ui.empty),emptyText=$('span',ui.empty);emptyTitle.textContent=q?'Sin coincidencias':'Busca o escanea un material';emptyText.textContent=q?'El código no pertenece al inventario cargado.':'La lista permanecerá vacía hasta que ingreses un código.';
    for(const group of groups)ui.results.appendChild(renderCard(group));updateSummary();
  }
  function recordNeedsLocation(record){return !!record&&String(recordValue(record,'count')).trim()!==''&&String(recordValue(record,'currentLocation')).trim()==='';}
  function showMissingLocation(record,next){state.pendingNavigation=next;state.missingRecordKey=record.key;ui.missingLocationMessage.innerHTML=`El material <strong>${escapeHtml(record.material)}</strong> tiene un conteo de <strong>${escapeHtml(formatValue(recordValue(record,'count')))}</strong>, pero no tiene ubicación.`;ui.missingLocationDialog.showModal();}
  function guardedNavigation(next){if(state.readOnly){next();return true;}const record=state.records.find(r=>r.key===state.viewingRecordKey);if(recordNeedsLocation(record)&&!state.missingLocationAcknowledged.has(record.key)){showMissingLocation(record,next);return false;}next();return true;}
  function handleSearchInput(){
    if(state.suppressSearchGuard){state.suppressSearchGuard=false;render();return;}
    const record=state.records.find(r=>r.key===state.searchRecordKey);if(record&&ui.searchInput.value!==state.searchBaseline&&recordNeedsLocation(record)&&!state.missingLocationAcknowledged.has(record.key)){
      const pending=ui.searchInput.value;ui.searchInput.value=state.searchBaseline;render();showMissingLocation(record,()=>{state.suppressSearchGuard=true;ui.searchInput.value=pending;state.searchBaseline=pending;render();ui.searchInput.focus();ui.searchInput.setSelectionRange?.(pending.length,pending.length);});return;
    }render();
  }
  function finalizeDictatedInput(){const compact=dictatedCodeValue(ui.searchInput.value),known=resolveKnownCode(compact);if(known||compact!==ui.searchInput.value)ui.searchInput.value=known||compact;render();}
  function normalizeLocationInput(value){return String(value??'').toUpperCase().replace(/\s/g,'').slice(0,10);}
  function sapValuesMatch(record,count,loc){return String(count).trim()!==''&&String(loc).trim()!==''&&countsMatch(record,count)&&normalizeLocationInput(loc)===normalizeLocationInput(record.sapLocations);}
  function renderCard(group){
    const material=group[0].material,record=activeRecordForGroup(group);state.activeLocations[material]=record.key;
    const count=recordValue(record,'count'),loc=recordValue(record,'currentLocation'),comment=recordValue(record,'comments'),keepManual=!!recordValue(record,'keepSap'),keepLocation=!!recordValue(record,'keepLocation'),keepAuto=!keepManual&&!keepLocation&&sapValuesMatch(record,count,loc),keep=keepManual||keepAuto,canAdd=!record.serialized&&!keepManual;
    const selected=state.selectedSeries[record.key]||[],seriesItems=recordSeriesItems(record),batchGroup=batchSplitGroup(group),materialStatus=record.serialized?'<span class="material-status serialized">Serializado</span>':batchGroup?'<span class="material-status batch">Batch</span>':'',issueSummary=recordIssueSummary(record),hasCriticalIssue=(record.seriesIssues||[]).some(issue=>issue.severity==='critical');
    const card=document.createElement('article');card.className='material-card';card.dataset.material=material;
    card.innerHTML=`<div class="card-main"><div class="material-line"><div><p class="material-code">MATERIAL · ${escapeHtml(material)}</p><h2 class="description">${escapeHtml(record.description||'Sin descripción')}</h2></div><button class="comment-button" type="button" aria-label="Comentario"><svg viewBox="0 0 48 48"><path d="M11 12H37Q40 12 40 15V31Q40 34 37 34H21L12 40V34H11Q8 34 8 31V15Q8 12 11 12Z"/><path d="M15 20H33M15 26H28"/></svg>${comment?'<span class="comment-dot"></span>':''}</button></div><div class="material-meta"><span>UM · ${escapeHtml(record.unitMeasure||'—')}</span>${materialStatus}<span>Valor unitario · $ ${escapeHtml(formatValue(record.unitValue))}</span></div></div>
      <div class="location-block"><span class="section-label">Storage Location (St Loc)</span><div class="loc-tabs">${group.map(r=>{const same=group.filter(x=>x.stloc===r.stloc),duplicate=same.length>1,type=stockTypeLabel(r.stockType),useBatch=batchGroup&&duplicate&&String(r.batch||'').trim(),fallback=duplicate?` · ${type||same.indexOf(r)+1}`:'',label=useBatch?`${escapeHtml(r.stloc)} ·<span class="loc-batch-value">${escapeHtml(String(r.batch).trim())}</span>`:`${escapeHtml(r.stloc)}${escapeHtml(fallback)}`;return `<button class="loc-tab ${useBatch?'batch-tab':''} ${r.key===record.key?'active':''}" data-key="${escapeHtml(r.key)}" type="button">${label}</button>`;}).join('')}</div></div>
      <div class="sap-strip"><div><span>SAP qty</span><strong>${escapeHtml(formatValue(record.sapQty))}</strong></div><div><span>SAP location</span><strong>${escapeHtml(record.sapLocations||'—')}</strong></div></div>
      <div class="entry-grid"><div class="entry-column"><div class="entry-field"><label>Conteo físico</label><div class="count-control"><input class="count-input" inputmode="decimal" value="${escapeHtml(count)}" ${keepManual||record.serialized?'disabled':''}>${canAdd?`<button class="add-count-button ${String(count).trim()===''?'hidden':''}" type="button" aria-label="Sumar otra cantidad">+</button>`:''}</div></div><label class="keep-option ${keepAuto?'auto-match':''}"><input class="keep-both-input" type="checkbox" ${keep?'checked':''}> <span>Mantener cantidad y ubicación SAP</span></label></div><div class="entry-column"><div class="entry-field location"><label>Ubicación actual</label><input class="location-input" maxlength="10" value="${escapeHtml(loc)}" ${keepManual||keepLocation?'disabled':''}></div><label class="keep-option"><input class="keep-location-input" type="checkbox" ${keepLocation?'checked':''}> <span>Mantener solo ubicación</span></label></div></div>
      ${issueSummary?`<div class="material-series-warning ${hasCriticalIssue?'critical':''}"><strong>Revisar series</strong><span>${escapeHtml(issueSummary)}</span></div>`:''}
      ${record.serialized?`<div class="series-panel"><div class="series-head"><strong>Series</strong><span>${selected.length} seleccionadas</span></div><div class="series-list">${seriesItems.length?seriesItems.map(s=>{const key=Model.seriesSelectionKey(s),mismatch=(record.alternateSerials||[]).includes(s);return `<button type="button" class="series-chip ${s.available?'available':''} ${mismatch?'location-mismatch':''} ${selected.includes(key)?'selected':''}" data-series-key="${escapeHtml(key)}" ${s.available?'':'disabled'}>${escapeHtml(s.serial)}${mismatch?`<small>${escapeHtml(s.location)} → ${escapeHtml(record.stloc)}</small>`:''}</button>`;}).join(''):'<div class="series-empty-warning">No hay series disponibles para este material. El conteo existente se conserva sin cambios.</div>'}</div></div>`:''}`;
    $$('.loc-tab',card).forEach(b=>b.addEventListener('click',()=>{state.activeLocations[material]=b.dataset.key;render();}));
    const countInput=$('.count-input',card);countInput.addEventListener('input',e=>{const v=e.target.value.trim();if(v===''||(Number.isFinite(Number(v))&&Number(v)>=0)){setChange(record,'count',v);$('.add-count-button',card)?.classList.toggle('hidden',v==='');}});countInput.addEventListener('change',e=>{const v=e.target.value.trim();if(v!==''&&(!Number.isFinite(Number(v))||Number(v)<0)){toast('Ingresa una cantidad válida');e.target.value=recordValue(record,'count');return;}setChange(record,'count',v);render();});
    const locationInput=$('.location-input',card);locationInput.addEventListener('input',e=>{const value=normalizeLocationInput(e.target.value);if(e.target.value!==value)e.target.value=value;setChange(record,'currentLocation',value);});
    $('.add-count-button',card)?.addEventListener('click',()=>openAddCount(record));
    $('.keep-both-input',card).addEventListener('change',e=>{const checked=e.target.checked;setChange(record,'keepSap',checked);if(checked){setChange(record,'keepLocation',false);setChange(record,'count',record.sapQty);setChange(record,'currentLocation',normalizeLocationInput(record.sapLocations));if(record.serialized){const wanted=Math.max(0,Math.trunc(numericCell(record.sapQty)||0)),available=record.serials.filter(s=>s.available).map(Model.seriesSelectionKey);state.selectedSeries[record.key]=available.slice(0,wanted);queuePersist();}}render();});
    $('.keep-location-input',card).addEventListener('change',e=>{setChange(record,'keepLocation',e.target.checked);if(e.target.checked){setChange(record,'keepSap',false);setChange(record,'currentLocation',normalizeLocationInput(record.sapLocations));}render();});
    $('.comment-button',card).addEventListener('click',()=>openComment(record));
    $$('.series-chip',card).forEach(b=>b.addEventListener('click',()=>toggleSeries(record,b.dataset.seriesKey)));
    if(state.readOnly)$$('.count-input,.location-input,.add-count-button,.keep-both-input,.keep-location-input,.series-chip',card).forEach(control=>control.disabled=true);
    return card;
  }
  function openAddCount(record){state.addRecordKey=record.key;const current=Number(recordValue(record,'count'))||0;ui.addCountMaterial.textContent=`${record.material} · Conteo actual: ${formatValue(current)}`;ui.addCountCurrent.textContent=formatValue(current);ui.addCountValue.value='';ui.addCountTotal.textContent=formatValue(current);ui.addCountDialog.showModal();ui.addCountValue.focus({preventScroll:true});ui.addCountValue.select();requestAnimationFrame(()=>{ui.addCountValue.focus({preventScroll:true});ui.addCountValue.select();});}
  function updateAddTotal(){const record=state.records.find(r=>r.key===state.addRecordKey),current=Number(recordValue(record,'count'))||0,extra=Math.max(0,Number(ui.addCountValue.value)||0);ui.addCountTotal.textContent=formatValue(Math.round((current+extra)*1000000)/1000000);}
  function closeLocationMenu(){ui.locationMenu.classList.add('hidden');ui.locationTrigger.setAttribute('aria-expanded','false');}
  function openLocationMenu(){ui.locationSearch.value='';renderLocationSuggestions('');ui.locationMenu.classList.remove('hidden');ui.locationTrigger.setAttribute('aria-expanded','true');}
  function populateLocationSelect(){const previous=state.selectedSapLocation;state.sapLocationOptions=[...new Set(state.records.map(r=>String(r.sapLocations||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es',{numeric:true}));state.selectedSapLocation=state.sapLocationOptions.includes(previous)?previous:'';ui.locationTriggerText.textContent=state.selectedSapLocation||'Seleccionar ubicación SAP';ui.locationSearch.value='';renderLocationSuggestions('');closeLocationMenu();}
  function renderLocationSuggestions(query){const q=normalize(query),matches=state.sapLocationOptions.filter(location=>!q||normalize(location).includes(q)).slice(0,40);ui.locationSuggestions.innerHTML='';for(const location of matches){const pending=state.records.filter(r=>r.sapLocations===location&&String(recordValue(r,'count')).trim()==='').length,button=document.createElement('button');button.type='button';button.setAttribute('role','option');button.innerHTML=`<strong>${escapeHtml(location)}</strong><span>${pending?`${pending} pendiente${pending===1?'':'s'}`:'Completa'}</span>`;button.addEventListener('click',()=>{state.selectedSapLocation=location;ui.locationTriggerText.textContent=location;closeLocationMenu();renderLocationReview();});ui.locationSuggestions.appendChild(button);}if(!matches.length)ui.locationSuggestions.innerHTML='<div class="location-suggestions-empty">Sin coincidencias</div>';}
  function countsMatch(record,count){const a=Number(record.sapQty),b=Number(count);return String(record.sapQty).trim()!==''&&Number.isFinite(a)&&Number.isFinite(b)?a===b:String(record.sapQty).trim()===String(count).trim();}
  function locationReviewRow(record){
    const count=recordValue(record,'count'),loc=String(recordValue(record,'currentLocation')).trim(),pending=String(count).trim()==='',match=!pending&&countsMatch(record,count),relocated=loc&&normalizeLocationInput(loc)!==normalizeLocationInput(record.sapLocations),row=document.createElement('button');
    row.type='button';row.className='location-row';
    row.innerHTML=`<div class="location-material"><i class="location-dot ${pending?'status-pending':match?'status-match':'status-difference'}"></i><div class="location-material-copy"><div><strong>${escapeHtml(record.material)}</strong><span class="description-small">${escapeHtml(record.description||'Sin descripción')}</span>${!pending&&!loc?'<span class="missing-location-note">Falta ubicación</span>':''}</div>${relocated?`<span class="relocated"><small>Reubicado</small>→ ${escapeHtml(loc)}</span>`:''}</div></div><div class="location-sap">${escapeHtml(formatValue(record.sapQty))}</div><div class="location-count ${pending?'pending':match?'match':'difference'}">${pending?'Pendiente':escapeHtml(formatValue(count))}</div><b class="location-chevron">›</b>`;
    row.addEventListener('click',()=>guardedNavigation(()=>openRecord(record)));return row;
  }
  function appendLocationTable(records,title=''){
    if(title){const heading=document.createElement('div');heading.className='location-section-title';heading.textContent=title;ui.locationResults.appendChild(heading);}
    const head=document.createElement('div');head.className='location-table-head';head.innerHTML='<span>Material</span><span>SAP</span><span>Conteo</span><span></span>';ui.locationResults.appendChild(head);
    for(const record of records)ui.locationResults.appendChild(locationReviewRow(record));
  }
  function renderLocationReview(){
    const location=state.selectedSapLocation;if(!location){ui.locationProgress.classList.add('hidden');ui.locationResults.innerHTML='';ui.locationEmpty.classList.add('hidden');return;}const normalizedLocation=normalizeLocationInput(location),records=state.records.filter(r=>String(r.sapLocations).trim()===location),foundHere=state.records.filter(r=>normalizeLocationInput(recordValue(r,'currentLocation'))===normalizedLocation&&normalizeLocationInput(r.sapLocations)!==normalizedLocation),reviewed=records.filter(r=>String(recordValue(r,'count')).trim()!=='').length,total=records.length,percent=total?Math.round(reviewed/total*100):0;
    ui.locationProgress.classList.toggle('hidden',!location);ui.locationProgressLabel.textContent=`Ubicación · ${location||'—'}`;ui.locationProgressCount.textContent=`${reviewed} de ${total} revisados`;ui.locationProgressPercent.textContent=`${percent}%`;ui.locationProgressBar.style.width=`${percent}%`;ui.locationResults.innerHTML='';
    const hasRows=records.length>0||foundHere.length>0;ui.locationEmpty.classList.toggle('hidden',hasRows);if(!hasRows)return;
    if(records.length)appendLocationTable(records);
    if(foundHere.length)appendLocationTable(foundHere,'Encontrados aquí · SAP indica otra ubicación');
  }
  function numericCell(value){
    if(typeof value==='number')return Number.isFinite(value)?value:0;
    const raw=String(value??'').trim();if(!raw)return 0;const direct=Number(raw);if(Number.isFinite(direct))return direct;
    let clean=raw.replace(/[^\d,.-]/g,'');
    if(clean.includes(',')&&clean.includes('.'))clean=clean.lastIndexOf(',')>clean.lastIndexOf('.')?clean.replace(/\./g,'').replace(',','.'):clean.replace(/,/g,'');
    else if(clean.includes(','))clean=clean.replace(/\./g,'').replace(',','.');
    const parsed=Number(clean);return Number.isFinite(parsed)?parsed:0;
  }
  function formatMoney(value){const sign=value<0?'-':'';return `${sign}$ ${new Intl.NumberFormat('es-CL',{maximumFractionDigits:0}).format(Math.abs(Math.round(value)))}`;}
  function statusRow(record,value,label,valueClass=''){
    const row=document.createElement('button');row.type='button';row.className='status-row';row.innerHTML=`<div class="status-row-copy"><strong>${escapeHtml(record.material)}</strong><span>${escapeHtml(record.description||'Sin descripción')}</span></div><div class="status-row-value ${valueClass}"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div><b>›</b>`;row.addEventListener('click',()=>guardedNavigation(()=>openRecord(record)));return row;
  }
  function renderStatus(){
    renderSeriesIssues();
    renderConsistencyStatus();
    const total=state.records.length,reviewedRecords=state.records.filter(r=>String(recordValue(r,'count')).trim()!==''),reviewed=reviewedRecords.length,pending=total-reviewed;
    const matches=reviewedRecords.filter(r=>countsMatch(r,recordValue(r,'count'))).length,differences=reviewed-matches;
    const missing=reviewedRecords.filter(r=>String(recordValue(r,'currentLocation')).trim()===''),relocated=state.records.filter(r=>{const loc=String(recordValue(r,'currentLocation')).trim();return loc&&normalizeLocationInput(loc)!==normalizeLocationInput(r.sapLocations);});
    const sapValue=state.records.reduce((sum,r)=>sum+numericCell(r.sapQty)*numericCell(r.unitValue),0),countedValue=reviewedRecords.reduce((sum,r)=>sum+numericCell(recordValue(r,'count'))*numericCell(r.unitValue),0),differenceValue=countedValue-sapValue;
    const valuedPercent=sapValue?countedValue/sapValue*100:0,reviewedPercent=total?Math.round(reviewed/total*100):0,series=Object.values(state.selectedSeries).reduce((n,a)=>n+a.length,0),seriesUnconfirmed=reviewedRecords.filter(r=>r.serialized&&Math.abs((state.selectedSeries[r.key]||[]).length-numericCell(recordValue(r,'count')))>.000001);
    ui.statusSapValue.textContent=formatMoney(sapValue);ui.statusCountedValue.textContent=formatMoney(countedValue);ui.statusDifference.textContent=formatMoney(differenceValue);ui.statusDifference.classList.toggle('difference',Math.abs(differenceValue)>.005);ui.statusPercent.textContent=`${new Intl.NumberFormat('es-CL',{minimumFractionDigits:2,maximumFractionDigits:2}).format(valuedPercent)} %`;ui.statusPercent.classList.toggle('difference',Math.abs(valuedPercent-100)>.05);
    ui.statusReviewed.textContent=`${reviewed} de ${total}`;ui.statusReviewedPercent.textContent=`${reviewedPercent}%`;ui.statusProgressBar.style.width=`${reviewedPercent}%`;ui.statusMatches.textContent=matches;ui.statusDifferences.textContent=differences;ui.statusPending.textContent=pending;ui.statusMissing.textContent=missing.length;ui.statusRelocated.textContent=relocated.length;ui.statusSeries.textContent=series;ui.statusMissingPanelCount.textContent=missing.length;ui.statusSeriesUnconfirmed.textContent=seriesUnconfirmed.length;ui.statusSeriesUnconfirmedPanelCount.textContent=seriesUnconfirmed.length;ui.statusSeriesUnconfirmedToggle.classList.toggle('hidden',!seriesUnconfirmed.length);if(!seriesUnconfirmed.length)ui.statusSeriesUnconfirmedPanel.classList.add('hidden');
    ui.statusSavedAt.textContent=saveStatusLabel();
    ui.statusMissingRows.innerHTML='';if(!missing.length)ui.statusMissingRows.innerHTML='<div class="status-list-empty">Todos los materiales contados tienen ubicación.</div>';else for(const r of missing)ui.statusMissingRows.appendChild(statusRow(r,formatValue(recordValue(r,'count')),'Conteo','negative'));
    ui.statusSeriesUnconfirmedRows.innerHTML='';for(const r of seriesUnconfirmed){const selected=(state.selectedSeries[r.key]||[]).length,count=formatValue(recordValue(r,'count'));ui.statusSeriesUnconfirmedRows.appendChild(statusRow(r,`${selected} de ${count}`,'Series / conteo','negative'));}
    const top=reviewedRecords.map(r=>({record:r,value:(numericCell(recordValue(r,'count'))-numericCell(r.sapQty))*numericCell(r.unitValue)})).filter(x=>Math.abs(x.value)>.005).sort((a,b)=>Math.abs(b.value)-Math.abs(a.value)).slice(0,5);
    ui.statusTopDifferences.innerHTML='';if(!top.length)ui.statusTopDifferences.innerHTML='<div class="status-list-empty">Sin diferencias valorizadas en las filas revisadas.</div>';else for(const item of top)ui.statusTopDifferences.appendChild(statusRow(item.record,formatMoney(item.value),'Diferencia valorizada',item.value<0?'negative':'positive'));
  }
  function showInventoryView(){hideMainViews();state.currentView='inventory';ui.inventoryView.classList.remove('hidden');setActiveNav('showInventory');}
  function showLocationsView(){hideMainViews();state.currentView='locations';ui.locationReviewView.classList.remove('hidden');setActiveNav('showLocations');populateLocationSelect();renderLocationReview();}
  function showStatusView(){hideMainViews();state.currentView='status';ui.statusView.classList.remove('hidden');setActiveNav('showStatus');renderStatus();}
  function showFinishView(){hideMainViews();state.currentView='finish';ui.finishView.classList.remove('hidden');setActiveNav('finishButton');prepareFinishView();}
  function openRecord(record){showInventoryView();state.activeLocations[record.material]=record.key;ui.searchInput.value=record.material;state.searchBaseline=record.material;render();window.scrollTo({top:0,behavior:'smooth'});}
  function focusRecordLocation(record){if(!record)return;openRecord(record);setTimeout(()=>{$('.location-input',ui.results)?.focus();},80);}
  function openComment(record){state.commentKey=record.key;ui.commentTitle.textContent=`${record.material} · ${record.stloc}`;ui.commentText.value=String(recordValue(record,'comments')||'').toUpperCase();ui.commentText.readOnly=state.readOnly;$('#saveComment').classList.toggle('hidden',state.readOnly);ui.commentDialog.showModal();}
  function toggleSeries(record,selectionKey){
    if(state.readOnly){toast('El inventario histórico es de solo lectura');return;}
    const previous=state.selectedSeries,next=Model.toggleSeriesSelection(state.records,previous,record,selectionKey);state.selectedSeries=next;
    for(const item of state.records){if(normalize(item.material)!==normalize(record.material))continue;const before=previous[item.key]||[],after=next[item.key]||[];if(before.length!==after.length||before.some((value,index)=>value!==after[index])){state.changes[item.key]={...(state.changes[item.key]||{}),count:after.length};state.missingLocationAcknowledged.delete(item.key);}}
    updateSummary();queuePersist();if(state.currentView==='locations')renderLocationReview();if(state.currentView==='status')renderStatus();render();
  }
  function updateSummary(){const rows=state.records.filter(r=>String(recordValue(r,'count')).trim()!==''||String(recordValue(r,'currentLocation')).trim()!==''||String(recordValue(r,'comments')).trim()!=='').length;ui.countedTotal.textContent=rows;ui.finishRows.textContent=rows;ui.finishSeries.textContent=Object.values(state.selectedSeries).reduce((n,a)=>n+a.length,0);}

  function blockUnknown(){const q=ui.searchInput.value.trim();if(q&&state.filtered.length===0){navigator.vibrate?.(120);toast(`Código ${q} bloqueado: no existe en el inventario`);}}

  function prepareAudio(){try{state.audioContext??=new (window.AudioContext||window.webkitAudioContext)();state.audioContext.resume?.();}catch{}}
  function playScanFeedback(){
    navigator.vibrate?.(55);const ctx=state.audioContext;if(!ctx)return;
    try{const oscillator=ctx.createOscillator(),gain=ctx.createGain(),now=ctx.currentTime;oscillator.type='sine';oscillator.frequency.setValueAtTime(1046,now);gain.gain.setValueAtTime(.09,now);gain.gain.exponentialRampToValueAtTime(.001,now+.14);oscillator.connect(gain);gain.connect(ctx.destination);oscillator.start(now);oscillator.stop(now+.15);}catch{}
  }

  function acceptScannedCode(value){
    const code=resolveKnownCode(value,{ocr:true})||String(value??'').trim();if(!code)return;stopScanner();guardedNavigation(()=>{showInventoryView();ui.searchInput.value=code;state.searchBaseline=code;playScanFeedback();render();blockUnknown();});
  }

  function isKnownCode(code){return !!resolveKnownCode(code,{ocr:true});}
  function editDistance(a,b){const previous=Array.from({length:b.length+1},(_,i)=>i);for(let i=1;i<=a.length;i++){let diagonal=previous[0];previous[0]=i;for(let j=1;j<=b.length;j++){const above=previous[j],cost=a[i-1]===b[j-1]?0:1;previous[j]=Math.min(previous[j]+1,previous[j-1]+1,diagonal+cost);diagonal=above;}}return previous[b.length];}
  function resolveOcrCode(value){
    const exact=resolveKnownCode(value,{ocr:true});if(exact)return exact;const compact=compactCode(value);if(compact.length<5||!/[0-9]/.test(compact))return '';
    const ranked=[...state.compactCodes.entries()].filter(([code])=>Math.abs(code.length-compact.length)<=1).map(([code,values])=>({code,values,distance:editDistance(compact,code)})).sort((a,b)=>a.distance-b.distance);
    if(!ranked.length)return '';const allowed=compact.length>=8?2:1,best=ranked[0],second=ranked[1];return best.distance<=allowed&&best.values.length===1&&(!second||second.distance>best.distance)?best.values[0]:'';
  }
  function considerScannedCode(value){
    const code=String(value??'').trim(),now=Date.now();if(!code)return;
    if(!isKnownCode(code)){ui.scannerHelp.textContent=`Lectura ${code} descartada: no existe en el Excel. Sigue apuntando.`;return;}
    if(state.scanCandidate===code&&now-state.scanCandidateAt<4000)state.scanHits+=1;else{state.scanCandidate=code;state.scanHits=1;}
    state.scanCandidateAt=now;
    if(state.scanHits<2){ui.scannerHelp.textContent=`Código ${code} detectado. Confirmando lectura…`;return;}
    acceptScannedCode(code);
  }

  async function prepareOcrWorker(){
    if(state.ocrWorker)return state.ocrWorker;if(state.ocrWorkerPromise)return state.ocrWorkerPromise;
    const base=new URL('./vendor/tesseract/',location.href).href,statusLabels={'loading tesseract core':'Cargando motor','initializing tesseract':'Iniciando motor','loading language traineddata':'Cargando reconocimiento','initializing api':'Preparando reconocimiento','recognizing text':'Leyendo captura'};
    state.ocrWorkerPromise=(async()=>{const worker=await window.Tesseract.createWorker('eng',window.Tesseract.OEM.LSTM_ONLY,{workerPath:`${base}worker.min.js`,corePath:`${base}core`,langPath:`${base}lang-fast`,cachePath:'inventario-ocr-v8-fast',workerBlobURL:false,logger:m=>{if(state.ocrBusy){const label=statusLabels[m.status]||'Preparando lector';ui.scannerHelp.textContent=`${label}… ${Math.round((m.progress||0)*100)}%`;}}});await worker.setParameters({tessedit_pageseg_mode:window.Tesseract.PSM.AUTO});state.ocrWorker=worker;return worker;})();
    try{return await state.ocrWorkerPromise;}catch(e){state.ocrWorkerPromise=null;throw e;}
  }

  async function captureAndRead(){
    if(state.ocrBusy||!ui.cameraVideo.videoWidth)return;
    if(!window.Tesseract){toast('El lector local no pudo cargarse.');return;}
    state.ocrBusy=true;ui.captureButton.disabled=true;ui.captureButton.classList.add('processing');ui.cameraFrame.classList.add('captured');ui.scannerHelp.textContent=state.ocrWorker?'Analizando captura…':'Preparando lector local…';
    try{
      const video=ui.cameraVideo,canvas=ui.captureCanvas;canvas.width=video.videoWidth;canvas.height=video.videoHeight;canvas.getContext('2d',{willReadFrequently:true}).drawImage(video,0,0,canvas.width,canvas.height);
      const findMatch=text=>{const candidates=[];for(const line of String(text||'').toUpperCase().split(/\r?\n/)){candidates.push(...(line.match(/[A-Z0-9][A-Z0-9-]{3,}/g)||[]));const compact=line.replace(/[^A-Z0-9]/g,'');if(compact.length>=4)candidates.push(compact);}const match=candidates.map(resolveOcrCode).find(Boolean)||'';return {match,candidates};};
      const frameAspect=ui.cameraFrame.clientWidth/ui.cameraFrame.clientHeight,videoAspect=canvas.width/canvas.height;let viewX=0,viewY=0,viewW=canvas.width,viewH=canvas.height;
      if(videoAspect>frameAspect){viewW=canvas.height*frameAspect;viewX=(canvas.width-viewW)/2;}else if(videoAspect<frameAspect){viewH=canvas.width/frameAspect;viewY=(canvas.height-viewH)/2;}
      const makeCrop=(top,height,threshold=0)=>{const sx=Math.round(viewX+viewW*.08),sy=Math.round(viewY+viewH*top),sw=Math.round(viewW*.84),sh=Math.round(viewH*height),scale=Math.min(4,2000/sw),crop=document.createElement('canvas');crop.width=Math.round(sw*scale);crop.height=Math.round(sh*scale);const ctx=crop.getContext('2d',{willReadFrequently:true});ctx.filter='grayscale(1) contrast(1.32)';ctx.drawImage(canvas,sx,sy,sw,sh,0,0,crop.width,crop.height);ctx.filter='none';if(threshold){const image=ctx.getImageData(0,0,crop.width,crop.height),data=image.data;for(let i=0;i<data.length;i+=4){const value=data[i]<threshold?0:255;data[i]=data[i+1]=data[i+2]=value;}ctx.putImageData(image,0,0);}return crop;};
      await prepareOcrWorker();
      let seen=[],match='';const passes=[
        {crop:makeCrop(.455,.115),whitelist:'0123456789'},
        {crop:makeCrop(.445,.135,178),whitelist:'0123456789'},
        {crop:makeCrop(.44,.15),whitelist:'0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-'}
      ];
      for(const pass of passes){await state.ocrWorker.setParameters({tessedit_pageseg_mode:window.Tesseract.PSM.SINGLE_LINE,tessedit_char_whitelist:pass.whitelist,preserve_interword_spaces:'0'});const {data}=await state.ocrWorker.recognize(pass.crop),result=findMatch(data.text);seen.push(...result.candidates);if(result.match){match=result.match;break;}}
      if(!ui.scanDialog.open)return;if(match){acceptScannedCode(match);return;}
      const detected=[...new Set(seen.filter(v=>/\d/.test(v)))].slice(0,2).join(', ');ui.scannerHelp.textContent=detected?`Se leyó ${detected}, pero no coincide con el Excel. Volviendo a la cámara…`:'No se distinguió el código impreso. Volviendo a la cámara…';navigator.vibrate?.(80);await new Promise(resolve=>setTimeout(resolve,1600));
    }catch(e){ui.scannerHelp.textContent='No fue posible analizar la captura. Volviendo a la cámara…';await new Promise(resolve=>setTimeout(resolve,1200));}
    finally{state.ocrBusy=false;ui.captureButton.disabled=false;ui.captureButton.classList.remove('processing');ui.cameraFrame.classList.remove('captured');const ctx=ui.captureCanvas.getContext('2d');ctx?.clearRect(0,0,ui.captureCanvas.width,ui.captureCanvas.height);ui.captureCanvas.width=0;ui.captureCanvas.height=0;if(ui.scanDialog.open)ui.scannerHelp.textContent='Centra el código en la línea y toca el botón circular.';}
  }

  async function startScanner(){
    prepareAudio();
    try{await screen.orientation?.lock?.('portrait-primary');}catch{}
    if(!navigator.mediaDevices?.getUserMedia){toast('Este navegador no permite usar la cámara. Usa Safari o la búsqueda manual.');return;}
    ui.scannerHelp.textContent='Centra el código en la línea y toca el botón circular.';
    const cameraConstraints={video:{facingMode:{ideal:'environment'},width:{ideal:1920},height:{ideal:1080},frameRate:{ideal:30,max:30}}};
    try{
      state.stream=await navigator.mediaDevices.getUserMedia(cameraConstraints);ui.cameraVideo.srcObject=state.stream;await ui.cameraVideo.play();ui.scanDialog.showModal();
      const track=state.stream.getVideoTracks()[0],capabilities=track?.getCapabilities?.(),cameraTuning={};if(capabilities?.focusMode?.includes('continuous'))cameraTuning.focusMode='continuous';if(capabilities?.zoom?.max>1)cameraTuning.zoom=Math.min(capabilities.zoom.max,Math.max(capabilities.zoom.min,2));if(Object.keys(cameraTuning).length)track.applyConstraints({advanced:[cameraTuning]}).catch(()=>{});
    }catch(e){toast('No fue posible abrir la cámara. Revisa permisos o usa búsqueda manual.');}
  }
  function stopScanner(){if(state.scanTimer)cancelAnimationFrame(state.scanTimer);state.scanTimer=null;state.scanControls?.stop?.();state.scanControls=null;state.stream?.getTracks().forEach(t=>t.stop());state.stream=null;state.ocrBusy=false;ui.cameraVideo.srcObject=null;ui.captureButton.disabled=false;ui.captureButton.classList.remove('processing');ui.cameraFrame.classList.remove('captured');ui.captureCanvas.getContext('2d')?.clearRect(0,0,ui.captureCanvas.width,ui.captureCanvas.height);ui.captureCanvas.width=0;ui.captureCanvas.height=0;if(ui.scanDialog.open)ui.scanDialog.close();}

  async function buildExport(outputFileName=''){
    if(state.demo)throw new Error('El modo demostración no genera un Excel. Carga el libro real.');
    return Xlsx.buildExport({workbook:state.workbook,records:state.records,changes:state.changes,selectedSeries:state.selectedSeries,model:Model,outputFileName,FileType:File});
  }
  async function getExport(){if(criticalSeriesIssues().length&&!ui.issuesReviewCheck.checked)throw new Error('Confirma que entiendes las incidencias críticas antes de exportar.');if(!state.inventory||state.inventory.status!=='open')throw new Error('No existe un inventario activo para exportar.');await queuePersist();const metadata=Lifecycle.nextExportMetadata(state.inventory);ui.exportMessage.textContent='Preparando copia actualizada…';const file=await buildExport(metadata.fileName);ui.exportMessage.textContent=`Listo: ${file.name}`;return {file,metadata};}
  async function confirmActiveExport(file,metadata){const current=await db.getInventory(state.inventory?.id);if(!current||current.status!=='open')throw new Error('El inventario ya no está disponible para registrar la exportación.');const next={...current,exportRevision:metadata.revision,lastExportDate:metadata.date,lastExportFileName:metadata.fileName};await db.saveInventory(next);state.inventory=next;state.lastExport=file;state.lastExportMetadata=metadata;return next;}
  function download(file){const a=document.createElement('a');a.href=URL.createObjectURL(file);a.download=file.name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),10000);}
  async function downloadActiveExport(){const {file,metadata}=await getExport();download(file);await confirmActiveExport(file,metadata);toast(`Exportación ${metadata.revision} guardada.`);}
  async function shareExport(){try{const {file,metadata}=await getExport();if(navigator.canShare?.({files:[file]})){try{await navigator.share({files:[file],title:`${state.inventory?.name||'Inventario'} actualizado`,text:'Inventario actualizado al cierre de jornada.'});}catch(error){if(error?.name==='AbortError'){ui.exportMessage.textContent='Exportación cancelada.';return;}throw error;}}else{download(file);toast('Se descargó la copia. Guárdala en Archivos o adjúntala en Outlook.');}await confirmActiveExport(file,metadata);}catch(e){ui.exportMessage.textContent=e.message;toast(e.message);}}

  async function prepareExcelUpdate(file){
    if(!/\.xlsx$/i.test(file.name))throw new Error('Selecciona un archivo .xlsx.');
    const setting=await db.getSetting('activeInventoryId'),inventory=setting?.value?await db.getInventory(setting.value):null;if(!inventory||inventory.status!=='open')throw new Error('No existe un inventario activo para actualizar.');
    const currentWorkbook=state.inventory?.id===inventory.id?state.workbook:await loadWorkbook(inventory.bytes,inventory.fileName),bytes=await file.arrayBuffer(),candidateWorkbook=await loadWorkbook(bytes,file.name),candidateFingerprint=await inventorySessionId(candidateWorkbook);
    const plan=Lifecycle.buildExcelUpdatePlan({inventory,currentWorkbook,candidateWorkbook,candidateFingerprint,model:Model});state.pendingUpdate={inventory,currentWorkbook,candidateWorkbook,plan};
    const c=plan.counts;ui.updatePreviewSummary.innerHTML=`<div><span>Cantidades</span><strong>${c.quantities}</strong></div><div><span>Comentarios</span><strong>${c.comments}</strong></div><div><span>Ubicaciones</span><strong>${c.locations}</strong></div><div><span>Series</span><strong>${c.series}</strong></div>`;
    const detailValue=value=>Array.isArray(value)?(value.join(', ')||'—'):(String(value??'').trim()||'—');ui.updatePreviewDetails.innerHTML=plan.hasChanges?plan.details.map(detail=>`<article><div><strong>${escapeHtml(detail.material)}</strong><span>${escapeHtml(detail.label)}</span></div><div><small>Actual</small><b>${escapeHtml(detailValue(detail.before))}</b></div><div><small>Nuevo</small><b>${escapeHtml(detailValue(detail.after))}</b></div></article>`).join(''):'<div class="update-no-changes">No se detectaron cambios</div>';
    ui.updateWarnings.classList.toggle('hidden',!plan.unknownSeriesStyles.length);ui.updateWarnings.innerHTML=plan.unknownSeriesStyles.length?`<strong>${plan.unknownSeriesStyles.length} estilo(s) de serie no reconocido(s)</strong><span>Esas series conservarán su estado actual y no se interpretarán como desmarcadas.</span>`:'';
    $('#confirmExcelUpdate').disabled=!plan.hasChanges;ui.updatePreviewDialog.showModal();
  }

  async function applyPendingUpdate(){
    const pending=state.pendingUpdate;if(!pending)return;const backup=Lifecycle.createBackup(pending.inventory,'before-excel-update'),next=Lifecycle.applyExcelUpdate(pending.inventory,pending.plan);await db.saveInventory(next,{backup,activate:true});state.pendingUpdate=null;ui.updateFileInput.value='';await loadInventoryRecord(next,false);toast('Inventario actualizado. Respaldo anterior disponible.');
  }

  async function restoreLatestBackup(){
    const setting=await db.getSetting('activeInventoryId'),inventory=setting?.value?await db.getInventory(setting.value):null;if(!inventory)throw new Error('No existe un inventario activo.');const backups=(await db.backupsForInventory(inventory.id)).sort((a,b)=>(b.createdAt||0)-(a.createdAt||0)),latest=backups[0];if(!latest?.snapshot)throw new Error('No existe un respaldo recuperable.');if(!confirm(`Se restaurará el respaldo del ${new Intl.DateTimeFormat('es-CL',{dateStyle:'medium',timeStyle:'short'}).format(new Date(latest.createdAt))}. ¿Continuar?`))return;const safety=Lifecycle.createBackup(inventory,'before-backup-restore'),restored={...latest.snapshot,id:inventory.id,status:'open',closedAt:null,updatedAt:Date.now()};await db.saveInventory(restored,{backup:safety,activate:true});await loadInventoryRecord(restored,false);toast('Respaldo restaurado');
  }

  function showHome(){hideMainViews();ui.importView.classList.remove('hidden');ui.bottomNav.classList.add('hidden');state.currentView='home';state.readOnly=false;state.demo=false;state.newInventoryFormOpen=false;ui.demoExitButton.classList.add('hidden');$('#manageInventoriesButton').classList.remove('hidden');checkSavedSession();}
  async function showHistory(){
    hideMainViews();ui.historyView.classList.remove('hidden');ui.bottomNav.classList.add('hidden');state.currentView='history';ui.manageHistory.textContent=state.historyAdminMode?'Terminar':'Administrar';ui.manageHistory.setAttribute('aria-pressed',String(state.historyAdminMode));const setting=await db.getSetting('activeInventoryId'),all=(await db.allInventories()).filter(item=>item.id!==setting?.value),query=normalize(ui.historySearch.value),visible=all.filter(item=>!query||normalize(item.name).includes(query));ui.historyList.innerHTML='';
    for(const inventory of visible){const closed=inventory.status==='closed',item=document.createElement('article'),openButton=document.createElement('button');item.className='history-item';openButton.type='button';openButton.className='history-open-button';openButton.innerHTML=`<span><strong>${escapeHtml(inventory.name)}</strong><small>${closed?`Cerrado ${inventory.closedAt?new Intl.DateTimeFormat('es-CL',{dateStyle:'medium'}).format(new Date(inventory.closedAt)):'sin fecha registrada'}`:'Abierto · no activo'}</small></span><b>${closed?'Consultar':'Activar'} ›</b>`;openButton.addEventListener('click',async()=>{const active=await db.getSetting('activeInventoryId'),activeInventory=active?.value?await db.getInventory(active.value):null,lastView=activeInventory?.lastView||{};state.returnContext={activeInventoryId:active?.value||null,view:lastView.view||'inventory',search:lastView.search||'',scrollY:lastView.scrollY||0};if(!closed){if(!confirm(`“${inventory.name}” pasará a ser el inventario activo. El inventario actual permanecerá intacto. ¿Continuar?`))return;await db.putSetting('activeInventoryId',inventory.id);}state.historyAdminMode=false;await loadInventoryRecord(inventory,closed);});item.appendChild(openButton);if(closed&&state.historyAdminMode){const actions=document.createElement('div'),exportButton=document.createElement('button'),deleteButton=document.createElement('button');actions.className='history-item-actions';exportButton.type='button';exportButton.className='history-export-button';exportButton.textContent='Exportar Excel';exportButton.addEventListener('click',()=>exportClosedInventory(inventory).catch(error=>toast(error.message)));deleteButton.type='button';deleteButton.className='history-delete-button';deleteButton.textContent='Eliminar inventario';deleteButton.setAttribute('aria-label',`Eliminar inventario ${inventory.name}`);deleteButton.addEventListener('click',()=>openDeleteInventoryDialog(inventory));actions.append(exportButton,deleteButton);item.appendChild(actions);}ui.historyList.appendChild(item);}
    ui.historyEmpty.classList.toggle('hidden',visible.length>0);
  }
  function openDeleteInventoryDialog(inventory){state.pendingDeleteInventory={id:inventory.id,name:inventory.name};ui.deleteInventoryName.textContent=`“${inventory.name}”`;ui.deleteInventoryCheck.checked=false;ui.confirmDeleteInventory.disabled=true;ui.deleteInventoryDialog.showModal();}
  async function deletePendingInventory(){const pending=state.pendingDeleteInventory;if(!pending||!ui.deleteInventoryCheck.checked)return;ui.confirmDeleteInventory.disabled=true;try{await db.deleteClosedInventory(pending.id);state.pendingDeleteInventory=null;ui.deleteInventoryDialog.close('deleted');await showHistory();toast(`Inventario “${pending.name}” eliminado.`);}catch(error){ui.confirmDeleteInventory.disabled=false;toast(error.message);}}
  async function exportClosedInventory(inventory){const metadata=Lifecycle.closedExportMetadata(inventory);if(!metadata)throw new Error('Este inventario histórico no conserva una revisión y fecha verificables para exportar.');let file;if(inventory.finalBackupBytes)file=new File([inventory.finalBackupBytes],metadata.fileName,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});else{const workbook=await loadWorkbook(inventory.bytes,inventory.fileName);file=await Xlsx.buildExport({workbook,records:workbook.records,changes:inventory.changes||{},selectedSeries:inventory.selectedSeries||{},model:Model,outputFileName:metadata.fileName,FileType:File});}download(file);toast(`Exportado ${metadata.fileName}`);}
  async function returnToActive(){const setting=await db.getSetting('activeInventoryId'),active=setting?.value?await db.getInventory(setting.value):null;if(active&&active.status==='open'){await loadInventoryRecord(active,false);const view=state.returnContext?.view;if(view==='locations')showLocationsView();else if(view==='status')showStatusView();window.scrollTo(0,state.returnContext?.scrollY||0);}else showHome();}

  function prepareFinishView(){updateSummary();const critical=criticalSeriesIssues().length;ui.reviewCheck.checked=false;ui.issuesReviewCheck.checked=false;ui.finishIssuesWarning.classList.toggle('hidden',!critical);ui.issuesReviewLabel.classList.toggle('hidden',!critical);if(critical){ui.finishIssuesTitle.textContent=`${critical} incidencia${critical===1?'':'s'} crítica${critical===1?'':'s'} de series`;ui.finishIssuesText.textContent='Puedes exportar un avance parcial, pero las asociaciones indicadas siguen pendientes y no se corrigieron automáticamente.';}ui.exportMessage.textContent='Se genera una copia. El archivo original no se modifica.';ui.definitiveCloseSection.classList.toggle('hidden',state.readOnly||state.demo);updateFinishActions();}

  async function handleDefinitiveClose(){
    try{
      if(!state.inventory||state.readOnly)throw new Error('No existe un inventario abierto para cerrar.');
      if(!ui.reviewCheck.checked)throw new Error('Confirma primero que revisaste los registros de la jornada.');
      if(ui.confirmCloseInventory.dataset.stage!=='confirm'){
        const {file,metadata}=await getExport();state.lastExportBytes=await file.arrayBuffer();download(file);await confirmActiveExport(file,metadata);ui.confirmCloseInventory.dataset.stage='confirm';ui.confirmCloseInventory.textContent='Confirmar cierre definitivo';toast('Respaldo generado. Verifica que quedó guardado y confirma el cierre.');return;
      }
      const backup=Lifecycle.createBackup(state.inventory,'before-definitive-close'),closed=Lifecycle.closeInventory({...state.inventory,changes:state.changes,selectedSeries:state.selectedSeries},{backupFileName:state.lastExport?.name,exportRevision:state.lastExportMetadata?.revision,exportDate:state.lastExportMetadata?.date,finalBackupBytes:state.lastExportBytes});await db.saveInventory(closed,{backup,clearActive:true});ui.closeInventoryDialog.close();ui.confirmCloseInventory.dataset.stage='';ui.confirmCloseInventory.textContent='Generar respaldo y cerrar';state.inventory=null;state.workbook=null;state.records=[];state.changes={};state.selectedSeries={};state.lastExport=null;state.lastExportMetadata=null;state.lastExportBytes=null;showHome();toast('Inventario cerrado y disponible en el historial.');
    }catch(error){toast(error.message);}
  }

  ui.fileInput.addEventListener('change',async e=>{const f=e.target.files[0];if(!f)return;try{await prepareNewInventory(f);}catch(err){ui.fileInput.value='';toast(err.message);}});
  ui.updateFileInput.addEventListener('change',async e=>{const f=e.target.files[0];if(!f)return;try{await prepareExcelUpdate(f);}catch(err){ui.updateFileInput.value='';toast(err.message);}});
  $('#confirmCreateInventory').addEventListener('click',e=>{e.preventDefault();createPendingInventory().then(()=>ui.createInventoryDialog.close()).catch(error=>toast(error.message));});
  $('#confirmExcelUpdate').addEventListener('click',e=>{e.preventDefault();applyPendingUpdate().then(()=>ui.updatePreviewDialog.close()).catch(error=>toast(error.message));});
  ui.restoreButton.addEventListener('click',()=>restoreLatest().catch(e=>toast(e.message)));
  ui.startNewInventoryButton.addEventListener('click',()=>{state.newInventoryFormOpen=true;checkSavedSession();});$('#cancelNewInventory').addEventListener('click',()=>{state.newInventoryFormOpen=false;checkSavedSession();});
  ui.historyButton.addEventListener('click',()=>{state.historyAdminMode=false;showHistory().catch(e=>toast(e.message));});ui.manageHistory.addEventListener('click',()=>{state.historyAdminMode=!state.historyAdminMode;showHistory().catch(e=>toast(e.message));});$('#closeHistory').addEventListener('click',()=>{state.historyAdminMode=false;showHome();});ui.historySearch.addEventListener('input',()=>showHistory().catch(e=>toast(e.message)));$('#returnToActive').addEventListener('click',()=>returnToActive().catch(e=>toast(e.message)));
  ui.deleteInventoryCheck.addEventListener('change',()=>{ui.confirmDeleteInventory.disabled=!ui.deleteInventoryCheck.checked;});ui.confirmDeleteInventory.addEventListener('click',deletePendingInventory);ui.deleteInventoryDialog.addEventListener('close',()=>{state.pendingDeleteInventory=null;ui.deleteInventoryCheck.checked=false;ui.confirmDeleteInventory.disabled=true;});
  $('#manageInventoriesButton').addEventListener('click',()=>queuePersist().finally(showHome));
  ui.restoreBackupButton.addEventListener('click',()=>restoreLatestBackup().catch(e=>toast(e.message)));
  ui.retryRecoveryLookup.addEventListener('click',checkSavedSession);
  ui.demoButton.addEventListener('click',()=>{state.inventory=null;state.workbook=demoWorkbook();state.records=state.workbook.records;state.filtered=state.records;state.seriesIssues=[];state.consistencyReport=state.workbook.seriesConsistencyReport;state.demo=true;state.readOnly=false;ui.finishButton.disabled=false;ui.inventoryName.textContent='Demostración';rebuildCodeIndex();enterInventory();toast('Modo demostración aislado');});
  ui.demoExitButton.addEventListener('click',showHome);
  ui.searchInput.addEventListener('focus',()=>{state.searchBaseline=ui.searchInput.value;state.searchRecordKey=state.viewingRecordKey;if(ui.searchInput.value)setTimeout(()=>ui.searchInput.select(),0);});
  ui.searchInput.addEventListener('input',handleSearchInput);ui.searchInput.addEventListener('change',()=>{finalizeDictatedInput();blockUnknown();});ui.searchInput.addEventListener('blur',finalizeDictatedInput);ui.searchInput.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();finalizeDictatedInput();blockUnknown();}});
  ui.scanButton.addEventListener('click',startScanner);$('#closeScanner').addEventListener('click',stopScanner);ui.scanDialog.addEventListener('close',stopScanner);
  ui.captureButton.addEventListener('click',captureAndRead);ui.cameraFrame.addEventListener('click',captureAndRead);ui.cameraFrame.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();captureAndRead();}});
  $('#showInventory').addEventListener('click',showInventoryView);$('#showLocations').addEventListener('click',showLocationsView);$('#showStatus').addEventListener('click',showStatusView);ui.finishButton.addEventListener('click',showFinishView);ui.locationTrigger.addEventListener('click',()=>ui.locationMenu.classList.contains('hidden')?openLocationMenu():closeLocationMenu());ui.locationSearch.addEventListener('input',()=>renderLocationSuggestions(ui.locationSearch.value));ui.locationSearch.addEventListener('keydown',e=>{if(e.key==='Enter'){const first=$('button',ui.locationSuggestions);if(first){e.preventDefault();first.click();}}else if(e.key==='Escape')closeLocationMenu();});document.addEventListener('click',e=>{if(!ui.locationMenu.classList.contains('hidden')&&!e.target.closest('.location-combobox'))closeLocationMenu();});
  ui.statusMissingToggle.addEventListener('click',()=>{ui.statusMissingPanel.classList.toggle('hidden');if(!ui.statusMissingPanel.classList.contains('hidden'))ui.statusMissingPanel.scrollIntoView({behavior:'smooth',block:'nearest'});});
  ui.statusSeriesUnconfirmedToggle.addEventListener('click',()=>{ui.statusSeriesUnconfirmedPanel.classList.toggle('hidden');if(!ui.statusSeriesUnconfirmedPanel.classList.contains('hidden'))ui.statusSeriesUnconfirmedPanel.scrollIntoView({behavior:'smooth',block:'nearest'});});
  ui.statusSeriesIssuesToggle.addEventListener('click',()=>{ui.statusSeriesIssuesPanel.classList.toggle('hidden');if(!ui.statusSeriesIssuesPanel.classList.contains('hidden'))ui.statusSeriesIssuesPanel.scrollIntoView({behavior:'smooth',block:'nearest'});});
  ui.statusConsistencyToggle.addEventListener('click',()=>{ui.statusConsistencyPanel.classList.toggle('hidden');if(!ui.statusConsistencyPanel.classList.contains('hidden'))ui.statusConsistencyPanel.scrollIntoView({behavior:'smooth',block:'nearest'});});
  ui.seriesIssuesBanner.addEventListener('click',()=>{showStatusView();ui.statusSeriesIssuesPanel.classList.remove('hidden');ui.statusSeriesIssuesPanel.scrollIntoView({behavior:'smooth',block:'nearest'});});
  ui.addCountValue.addEventListener('input',updateAddTotal);$('#confirmAddCount').addEventListener('click',()=>{const record=state.records.find(r=>r.key===state.addRecordKey),extra=Number(ui.addCountValue.value);if(!record||!Number.isFinite(extra)||extra<0){toast('Ingresa una cantidad válida');return;}const total=Math.round(((Number(recordValue(record,'count'))||0)+extra)*1000000)/1000000;setChange(record,'count',String(total));ui.addCountDialog.close();render();});
  $('#completeLocationButton').addEventListener('click',()=>{const record=state.records.find(r=>r.key===state.missingRecordKey);state.pendingNavigation=null;ui.missingLocationDialog.close('complete');focusRecordLocation(record);});
  $('#continueWithoutLocation').addEventListener('click',()=>{const record=state.records.find(r=>r.key===state.missingRecordKey),next=state.pendingNavigation;if(record)state.missingLocationAcknowledged.add(record.key);state.pendingNavigation=null;ui.missingLocationDialog.close('continue');next?.();});
  ui.missingLocationDialog.addEventListener('close',()=>{if(state.pendingNavigation){const record=state.records.find(r=>r.key===state.missingRecordKey);state.pendingNavigation=null;focusRecordLocation(record);}});
  ui.commentText.addEventListener('input',()=>{const start=ui.commentText.selectionStart,end=ui.commentText.selectionEnd,upper=ui.commentText.value.toUpperCase();if(ui.commentText.value!==upper){ui.commentText.value=upper;ui.commentText.setSelectionRange(start,end);}});$('#saveComment').addEventListener('click',e=>{e.preventDefault();const r=state.records.find(x=>x.key===state.commentKey);if(r){setChange(r,'comments',ui.commentText.value.trim().toUpperCase());ui.commentDialog.close();render();}});
  function updateFinishActions(){const allowed=ui.reviewCheck.checked&&(!criticalSeriesIssues().length||ui.issuesReviewCheck.checked);ui.shareButton.disabled=!allowed;ui.downloadButton.disabled=!allowed;}
  ui.reviewCheck.addEventListener('change',updateFinishActions);ui.issuesReviewCheck.addEventListener('change',updateFinishActions);
  ui.shareButton.addEventListener('click',shareExport);ui.downloadButton.addEventListener('click',async()=>{try{await downloadActiveExport();}catch(e){ui.exportMessage.textContent=e.message;toast(e.message);}});
  ui.closeInventoryButton.addEventListener('click',()=>{ui.definitiveCloseCheck.checked=false;ui.confirmCloseInventory.disabled=true;ui.confirmCloseInventory.dataset.stage='';ui.confirmCloseInventory.textContent='Generar respaldo y cerrar';ui.closeInventoryDialog.showModal();});ui.definitiveCloseCheck.addEventListener('change',()=>{ui.confirmCloseInventory.disabled=!ui.definitiveCloseCheck.checked;});ui.confirmCloseInventory.addEventListener('click',handleDefinitiveClose);
  document.addEventListener('visibilitychange',()=>{if(document.hidden&&!state.demo)queuePersist();});window.addEventListener('pagehide',()=>{if(!state.demo)queuePersist();});
  window.addEventListener('online',()=>$('#networkStatus').textContent='Local');window.addEventListener('offline',()=>$('#networkStatus').textContent='Sin conexión');
  for(const dialog of $$('dialog'))dialog.addEventListener('close',()=>{if(dialog.contains(ui.toast))document.body.appendChild(ui.toast);});ui.updatePreviewDialog.addEventListener('close',()=>{state.pendingUpdate=null;ui.updateFileInput.value='';});
  checkSavedSession();
})();
