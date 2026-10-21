/* Client-side document triage. No file content is sent over the network. */
const input=document.querySelector('#fileInput'),drop=document.querySelector('#dropzone'),results=document.querySelector('#results'),tbody=document.querySelector('#fileRows'),summary=document.querySelector('#summary'),count=document.querySelector('#fileCount'),download=document.querySelector('#download');
if(window.pdfjsLib)pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
const types=[
 ['Финансовая отчетность',['бухгалтерск','баланс','отчет о финансов','форма 1','форма 2','финансовая отчетность']],
 ['Налоговая отчетность',['налог','ндс','деклараци','прибыл','усн','налоговая отчетность']],
 ['Банковские выписки и справки',['выписк','справк.*банк','bankreference','остаток на счет','движение средств']],
 ['Кредитная история (БКИ)',['бки','кредитн.*истор','кредитный отчет']],
 ['Анкета / сведения о заемщике',['анкета','опросник','сведения о заемщике','карточка предприятия']],
 ['Учредительные и регистрационные',['егрюл','устав','учредител','регистраци','лист записи','октэи']],
 ['Договоры и обязательства',['договор','контракт','соглашени','лизинг','гаранти','поручитель','кредитн.*соглаш']],
 ['Дебиторская / кредиторская задолженность',['карточка счета','карточка счёта','счет 60','счет 62','счёт 60','счёт 62','расшифровка задолженности']],
 ['Бизнес-план / прогноз',['тэо','бизнес-план','бизнес план','план доходов','прогноз','финансовая модель']],
 ['Документы поручителя / физлица',['паспорт','снилс','2ндфл','ндфл','справка о доходах','водительское удостоверение','поручител']],
 ['Имущество / обеспечение',['залог','оценка имущества','право собственности','свидетельство','обременени']],
 ['Переписка / пояснения',['письмо','пояснени','переписк','\.msg$']],
 ['Прочее',['.*']]
];
const pretty={'.pdf':'PDF','.xlsx':'Excel','.xls':'Excel','.docx':'Word','.doc':'Word','.xml':'XML','.jpg':'Изображение','.jpeg':'Изображение','.png':'Изображение','.tif':'Изображение','.tiff':'Изображение','.txt':'Текст','.csv':'CSV','.rtf':'RTF'};
let docs=[];
document.querySelector('#chooseFolder').onclick=()=>{input.setAttribute('webkitdirectory','');input.click()};
document.querySelector('#chooseFiles').onclick=()=>{input.removeAttribute('webkitdirectory');input.click()};
input.addEventListener('change',()=>{if(input.files.length)processFiles([...input.files]);input.value=''});
for(const ev of ['dragenter','dragover'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')});
for(const ev of ['dragleave','drop'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')});
drop.addEventListener('drop',e=>{const f=[...e.dataTransfer.files];if(f.length)processFiles(f)});
function classify(name,text){const s=(name+' '+text.slice(0,2500)).toLowerCase();for(const [label,patterns] of types.slice(0,-1))if(patterns.some(p=>new RegExp(p,'i').test(s)))return label;return 'Не определён'}
function find(pattern,text){const m=text.match(pattern);return m?m[1]||m[0]:''}
function fields(text){const clean=text.replace(/\s+/g,' ');return {ИНН:find(/\bИНН\s*[:№]?\s*(\d{10}|\d{12})\b/i,clean),КПП:find(/\bКПП\s*[:№]?\s*(\d{9})\b/i,clean),ОГРН:find(/\bОГРН(?:ИП)?\s*[:№]?\s*(\d{13}|\d{15})\b/i,clean),Период:find(/(?:за|на|период)\s*((?:19|20)\d{2}(?:\s*года?)?(?:\s*[-–—]\s*(?:19|20)?\d{2})?)/i,clean),Суммы: [...new Set((clean.match(/\b\d{1,3}(?:[\s\u00a0]\d{3})+(?:[,.]\d{1,2})?\s*(?:руб(?:лей|ля|\.)?|₽)?/gi)||[]).slice(0,8))].join('; '),Фрагмент:clean.slice(0,320)} }
async function readText(file){const ext=file.name.toLowerCase().split('.').pop();try{if(ext==='pdf'&&window.pdfjsLib){const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;let out='';for(let i=1;i<=Math.min(pdf.numPages,20);i++){const page=await pdf.getPage(i),c=await page.getTextContent();out+=c.items.map(x=>x.str).join(' ')+' ';}return out}
if(ext==='docx'&&window.mammoth){const r=await mammoth.extractRawText({arrayBuffer:await file.arrayBuffer()});return r.value}
if(ext==='xlsx'||ext==='xls'){const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellText:true});return wb.SheetNames.map(n=>XLSX.utils.sheet_to_csv(wb.Sheets[n])).join('\n').slice(0,100000)}
if(ext==='xml'||ext==='txt'||ext==='csv'||ext==='rtf')return (await file.text()).slice(0,100000);
return ''}catch(e){return ''}}
async function processFiles(files){results.classList.remove('hidden');docs=[];tbody.innerHTML='<tr><td colspan="4">Читаем файлы…</td></tr>';download.disabled=true;for(const file of files){const text=await readText(file);const type=classify((file.webkitRelativePath||'')+' '+file.name,text);const f=fields(text);docs.push({name:file.name,path:file.webkitRelativePath||file.name,size:file.size,format:pretty['.'+file.name.split('.').pop().toLowerCase()]||file.name.split('.').pop().toUpperCase(),type,text,fields:f,readable:!!text})}render();download.disabled=!window.XLSX}
function render(){count.textContent=`· ${docs.length} файл(ов)`;const groups={};docs.forEach(d=>groups[d.type]=(groups[d.type]||0)+1);summary.innerHTML=[['Файлов',docs.length],['Типов документов',Object.keys(groups).length],['С текстом',docs.filter(d=>d.readable).length],['Нужно проверить',docs.filter(d=>d.type==='Не определён'||!d.readable).length]].map(x=>`<div class="summary-card"><strong>${x[1]}</strong><span>${x[0]}</span></div>`).join('');tbody.innerHTML=docs.map(d=>{const ex=[d.fields.ИНН&&`ИНН ${d.fields.ИНН}`,d.fields.Период&&`Период ${d.fields.Период}`,d.fields.Суммы&&d.fields.Суммы.slice(0,75)].filter(Boolean).join(' · ')||'—';return `<tr><td title="${esc(d.path)}">${esc(d.path)}</td><td><span class="tag ${d.type==='Не определён'?'unknown':''}">${esc(d.type)}</span></td><td>${esc(ex)}</td><td class="status">${d.readable?'Текст доступен':'Текст не извлечён'}</td></tr>`}).join('');download.disabled=false}
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
download.onclick=()=>{if(!docs.length)return;const wb=XLSX.utils.book_new();const intro=[['РАЗБОР ДОКУМЕНТОВ ЗАЕМЩИКА'],['Сформировано',new Date().toLocaleString('ru-RU')],['Файлов в пакете',docs.length],['Назначение','Подготовка материалов для проверки специалистом; автоматическое кредитное решение не формируется.'],['Ограничение','Автоматически извлечённые значения требуют проверки по первоисточникам. Сканированные изображения требуют OCR.'],[],['ИНН','КПП','ОГРН','Файлов с извлеченным ИНН',docs.filter(d=>d.fields.ИНН).length]];const cover=XLSX.utils.aoa_to_sheet(intro);cover['!cols']=[{wch:32},{wch:42},{wch:22},{wch:28},{wch:18}];XLSX.utils.book_append_sheet(wb,cover,'Обзор');
const arr=[['Путь / файл','Формат','Определённый тип','Размер, байт','Извлечение текста','ИНН','КПП','ОГРН','Период (кандидат)','Суммы (кандидаты)','Фрагмент текста']];docs.forEach(d=>arr.push([d.path,d.format,d.type,d.size,d.readable?'Да':'Нет',d.fields.ИНН,d.fields.КПП,d.fields.ОГРН,d.fields.Период,d.fields.Суммы,d.fields.Фрагмент]));const ws=XLSX.utils.aoa_to_sheet(arr);ws['!cols']=[{wch:52},{wch:14},{wch:34},{wch:14},{wch:20},{wch:16},{wch:14},{wch:16},{wch:22},{wch:45},{wch:70}];ws['!autofilter']={ref:`A1:K${arr.length}`};ws['!freeze']={xSplit:0,ySplit:1};XLSX.utils.book_append_sheet(wb,ws,'Реестр документов');
const typesSummary=[['Тип документа','Количество'],...Object.entries(docs.reduce((o,d)=>(o[d.type]=(o[d.type]||0)+1,o),{})).sort((a,b)=>a[0].localeCompare(b[0],'ru'))];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(typesSummary),'Сводка по типам');
const checklist=[['Раздел проверки','Что проверить','Документы в пакете','Комментарий аналитика','Статус проверки'],...['Финансовая отчетность','Налоговая отчетность','Банковские выписки и справки','Кредитная история (БКИ)','Анкета / сведения о заемщике','Учредительные и регистрационные','Договоры и обязательства','Дебиторская / кредиторская задолженность','Бизнес-план / прогноз','Документы поручителя / физлица','Имущество / обеспечение'].map(t=>[t,'Сверить полноту, период, актуальность, реквизиты и данные с первоисточником',docs.filter(d=>d.type===t).map(d=>d.name).join('; '),'','Не проверено'])];const wc=XLSX.utils.aoa_to_sheet(checklist);wc['!cols']=[{wch:38},{wch:68},{wch:60},{wch:42},{wch:22}];wc['!autofilter']={ref:`A1:E${checklist.length}`};XLSX.utils.book_append_sheet(wb,wc,'Контрольный список');XLSX.writeFile(wb,`Разбор_заемщика_${new Date().toISOString().slice(0,10)}.xlsx`)};


