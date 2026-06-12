// Bewerbungsmanager - Dashboard + Intelligente Notizen
(() => {
  const STORAGE_KEY = 'bewerbungen_v1'
  const DEFAULT_STATUSES = ['Entwurf','Beworben','Eingangsbestätigung','Interview','Angebot','Absage']

  // DOM - Main elements
  const searchEl = document.getElementById('search')
  const listEl = document.getElementById('list')
  const newBtn = document.getElementById('newBtn')
  const importBtn = document.getElementById('importBtn')
  const modal = document.getElementById('modal')
  const form = document.getElementById('form')
  const formTitle = document.getElementById('formTitle')
  const cancelBtn = document.getElementById('cancel')
  const template = document.getElementById('appItem')
  const filterStatus = document.getElementById('filterStatus')
  const filterPriority = document.getElementById('filterPriority')
  const filterTag = document.getElementById('filterTag')
  const sortBy = document.getElementById('sortBy')

  const counts = {
    all: document.getElementById('countAll'),
    open: document.getElementById('countOpen'),
    interview: document.getElementById('countInterview'),
    offer: document.getElementById('countOffer'),
    reject: document.getElementById('countReject'),
    draft: document.getElementById('countDraft')
  }

  const remindersEl = document.getElementById('reminders')

  // DOM - Import Modal
  const importModal = document.getElementById('importModal')
  const importFile = document.getElementById('importFile')
  const importPreview = document.getElementById('importPreview')
  const confirmImport = document.getElementById('confirmImport')
  const cancelImport = document.getElementById('cancelImport')

  // DOM - Notes Modal
  const notesModal = document.getElementById('notesModal')
  const notesBtn = document.getElementById('notesBtn')
  const closeNotesBtn = document.getElementById('closeNotesBtn')
  const notesTabBtns = document.querySelectorAll('.notes-tab-btn')
  const templateBtns = document.querySelectorAll('.template-btn')
  const saveInterviewBtn = document.getElementById('saveInterview')
  const saveFeedbackBtn = document.getElementById('saveFeedback')
  const interviewText = document.getElementById('interviewText')
  const feedbackText = document.getElementById('feedbackText')
  const feedbackType = document.getElementById('feedbackType')

  // Charts
  let chartMonths, chartStatus, chartTrend

  // App state
  let apps = []
  let editingId = null
  let importedData = null

  function load(){
    try { apps = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]') } catch(e) { apps = [] }
  }

  function save(){
    localStorage.setItem(STORAGE_KEY, JSON.stringify(apps))
  }

  function id(){
    return Date.now().toString(36) + Math.random().toString(36).slice(2,6)
  }

  function pushHistory(a, text){
    a.history = a.history || []
    a.history.push({t:new Date().toISOString(), text})
  }

  function openModal(a){
    modal.classList.remove('hidden')
    formTitle.textContent = a ? 'Bewerbung bearbeiten' : 'Neue Bewerbung'
    editingId = a ? a.id : null
    form.company.value = a?.company || ''
    form.position.value = a?.position || ''
    form.date.value = a?.date || ''
    form.status.innerHTML = ''
    DEFAULT_STATUSES.forEach(s=>{
      const op = document.createElement('option')
      op.value = s
      op.textContent = s
      form.status.appendChild(op)
    })
    form.status.value = a?.status || DEFAULT_STATUSES[0]
    form.priority.value = a?.priority || ''
    form.tags.value = (a?.tags||[]).join(',')
    form.notes.value = a?.notes || ''
    form.reminder.value = a?.reminder || ''
  }

  function closeModal(){
    modal.classList.add('hidden')
    form.reset()
    editingId = null
  }

  function openImportModal(){
    importModal.classList.remove('hidden')
    importFile.value = ''
    importPreview.classList.add('hidden')
    confirmImport.classList.add('hidden')
    importedData = null
  }

  function closeImportModal(){
    importModal.classList.add('hidden')
    importedData = null
  }

  function openNotesModal(){
    notesModal.classList.remove('hidden')
    renderNotesUI()
  }

  function closeNotesModal(){
    notesModal.classList.add('hidden')
  }

  // Word Import Functions
  async function extractTextFromDocx(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result
          const zip = await JSZip.loadAsync(arrayBuffer)
          
          // Read document.xml
          const docXml = await zip.file('word/document.xml').async('string')
          const parser = new DOMParser()
          const xmlDoc = parser.parseFromString(docXml, 'application/xml')
          
          // Extract text from paragraphs
          const paragraphs = xmlDoc.getElementsByTagName('w:p')
          let text = ''
          let boldTexts = []
          
          for (let p of paragraphs) {
            let paragraphText = ''
            let isBold = false
            const runs = p.getElementsByTagName('w:r')
            
            for (let r of runs) {
              const rPr = r.getElementsByTagName('w:rPr')[0]
              if (rPr) {
                const bold = rPr.getElementsByTagName('w:b')[0]
                if (bold) isBold = true
              }
              
              const textElements = r.getElementsByTagName('w:t')
              for (let t of textElements) {
                paragraphText += t.textContent
              }
            }
            
            if (paragraphText) {
              text += paragraphText + '\n'
              if (isBold && boldTexts.length < 3) {
                boldTexts.push(paragraphText.trim())
              }
            }
          }
          
          resolve({ text, boldTexts })
        } catch (err) {
          reject(err)
        }
      }
      reader.onerror = reject
      reader.readAsArrayBuffer(file)
    })
  }

  function parseImportedData(text, boldTexts) {
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean)
    
    const data = {
      company: '',
      position: '',
      date: '',
      notes: text
    }
    
    // Extract company from bold text or first meaningful line
    if (boldTexts.length > 0) {
      data.company = boldTexts[0]
    }
    
    // Look for common patterns
    const companyKeywords = ['gmbh', 'ag', 'inc', 'ltd', 'group', 'company']
    for (let line of lines) {
      const lowerLine = line.toLowerCase()
      if (companyKeywords.some(kw => lowerLine.includes(kw))) {
        data.company = line
        break
      }
    }
    
    // Improved position detection - look for position keywords but be smarter about matches
    const positionKeywords = [
      'stelle', 'position', 'role', 'job', 'beruf', 'posten',
      'engineer', 'developer', 'designer', 'manager', 'analyst', 
      'praktikant', 'berater', 'sachbearbeiter', 'koordinator',
      'senior', 'junior', 'associate', 'specialist'
    ]
    
    // First pass: look for lines that START or STRONGLY match position keywords
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]
      const lowerLine = line.toLowerCase()
      
      // Skip very short or very long lines (likely not a position)
      if (line.length < 3 || line.length > 150) continue
      
      // Check if line STARTS with position keyword or has it prominently
      const startsWithKeyword = positionKeywords.some(kw => lowerLine.startsWith(kw))
      if (startsWithKeyword) {
        data.position = line
        break
      }
      
      // Check for position-like patterns (e.g., "Senior Developer", "Jr. Accountant")
      const hasPositionPattern = /^(senior|junior|jr|lead|head|chief|assistant|associate|trainee|intern|apprentice)\b/i.test(line)
      if (hasPositionPattern && !lowerLine.includes('manager') || lowerLine.includes('job')) {
        data.position = line
        break
      }
    }
    
    // Second pass: if not found, look for keywords anywhere in lines
    if (!data.position) {
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i]
        const lowerLine = line.toLowerCase()
        
        // Skip very short or very long lines
        if (line.length < 3 || line.length > 150) continue
        
        // Look for strong job title indicators
        if (positionKeywords.some(kw => lowerLine.includes(kw))) {
          // Avoid taking lines that are just keywords within paragraphs
          const wordCount = line.split(/\s+/).length
          if (wordCount <= 5) {
            data.position = line
            break
          }
        }
      }
    }
    
    // Look for dates (DD.MM.YYYY or YYYY-MM-DD format)
    const datePattern = /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{4}|\d{4}[.\/-]\d{1,2}[.\/-]\d{1,2})/g
    const dateMatches = text.match(datePattern)
    if (dateMatches) {
      data.date = dateMatches[0]
    }
    
    return data
  }

  function renderNotesUI(){
    const app = apps.find(a=>a.id===editingId)
    if(!app) return

    // Interview notes
    const interviewNotes = document.getElementById('interviewNotes')
    interviewNotes.innerHTML = ''
    (app.notes_interviews || []).forEach(note => {
      const div = document.createElement('div')
      div.className = 'note-item'
      div.innerHTML = `<strong>🎤 Interview</strong><br/><div class="date">${new Date(note.date).toLocaleString('de-DE')}</div><div>${note.content.replace(/\n/g,'<br>')}</div>`
      interviewNotes.appendChild(div)
    })

    // Feedback notes
    const feedbackNotes = document.getElementById('feedbackNotes')
    feedbackNotes.innerHTML = ''
    (app.notes_feedback || []).forEach(note => {
      const icons = {rejection:'❌',success:'✅',experience:'🎓'}
      const div = document.createElement('div')
      div.className = 'note-item'
      div.innerHTML = `<strong>${icons[note.type]} ${note.type==='rejection'?'Absage':note.type==='success'?'Erfolg':'Learnings'}</strong><br/><div class="date">${new Date(note.date).toLocaleString('de-DE')}</div><div>${note.content.replace(/\n/g,'<br>')}</div>`
      feedbackNotes.appendChild(div)
    })

    // History
    const history = document.getElementById('notesHistory')
    history.innerHTML = ''
    const allNotes = [
      ...(app.notes_interviews||[]).map(n=>({type:'interview',icon:'🎤',...n})),
      ...(app.notes_feedback||[]).map(n=>({type:'feedback',icon:n.type==='rejection'?'❌':n.type==='success'?'✅':'🎓',...n}))
    ].sort((a,b)=>new Date(b.date)-new Date(a.date))
    
    allNotes.forEach(note => {
      const div = document.createElement('div')
      div.className = 'note-item'
      div.innerHTML = `<strong>${note.icon}</strong><div class="date">${new Date(note.date).toLocaleString('de-DE')}</div><div>${note.content.substring(0,100)}...</div>`
      history.appendChild(div)
    })
  }

  function renderFilters(){
    filterStatus.innerHTML = '<option value="">Alle Status</option>'
    DEFAULT_STATUSES.forEach(s=>{
      const op = document.createElement('option')
      op.value = s
      op.textContent = s
      filterStatus.appendChild(op)
    })
  }

  function renderList(){
    const q = searchEl.value.trim().toLowerCase()
    const fs = filterStatus.value
    const fp = filterPriority.value
    const ft = filterTag.value.trim().toLowerCase()
    let data = apps.slice()
    
    if(q) data = data.filter(a=>[a.company,a.position,(a.tags||[]).join(' '),a.notes].join(' ').toLowerCase().includes(q))
    if(fs) data = data.filter(a=>a.status===fs)
    if(fp) data = data.filter(a=>a.priority===fp)
    if(ft) data = data.filter(a=>(a.tags||[]).map(t=>t.toLowerCase()).includes(ft))
    
    if(sortBy.value==='date_asc') data.sort((a,b)=>(a.date||'').localeCompare(b.date||''))
    else if(sortBy.value==='company') data.sort((a,b)=>(a.company||'').localeCompare(b.company||''))
    else data.sort((a,b)=>(b.date||'').localeCompare(a.date||''))

    listEl.innerHTML = ''
    data.forEach(a=>{
      const node = template.content.cloneNode(true)
      const li = node.querySelector('li')
      li.querySelector('.company').textContent = a.company || '(kein Unternehmen)'
      li.querySelector('.position').textContent = a.position || ''
      li.querySelector('.date').textContent = a.date || ''
      const s = li.querySelector('.status')
      s.textContent = a.status || ''
      if(a.priority==='Hoch') li.style.borderLeft='4px solid #e63946'
      li.querySelector('.edit').addEventListener('click',()=>{openModal(a)})
      li.querySelector('.del').addEventListener('click',()=>{
        if(confirm('Löschen?')){
          apps = apps.filter(x=>x.id!==a.id)
          pushHistory(a,'Gelöscht')
          save()
          renderAll()
        }
      })
      li.addEventListener('dblclick',()=>{openModal(a)})
      listEl.appendChild(node)
    })
    renderCounts()
    renderCharts()
    renderReminders()
  }

  function renderCounts(){
    counts.all.textContent = apps.length
    counts.open.textContent = apps.filter(a=>a.status && a.status.toLowerCase().includes('beworben')===false && a.status!=='Absage').length
    counts.interview.textContent = apps.filter(a=>a.status==='Interview').length
    counts.offer.textContent = apps.filter(a=>a.status==='Angebot').length
    counts.reject.textContent = apps.filter(a=>a.status==='Absage').length
    counts.draft.textContent = apps.filter(a=>a.status==='Entwurf').length
  }

  function renderCharts(){
    const months = {}
    const now = new Date()
    for(let i=5;i>=0;i--){
      const d=new Date(now.getFullYear(),now.getMonth()-i,1)
      months[d.toLocaleString('de-DE',{month:'short',year:'numeric'})]=0
    }
    apps.forEach(a=>{
      if(a.date){
        const d=new Date(a.date)
        const key=d.toLocaleString('de-DE',{month:'short',year:'numeric'})
        if(key in months) months[key]++
      }
    })
    const mLabels = Object.keys(months)
    const mData = Object.values(months)

    if(typeof Chart!=='undefined'){
      if(chartMonths) chartMonths.destroy()
      chartMonths = new Chart(document.getElementById('chartMonths'),{
        type:'bar',
        data:{labels:mLabels,datasets:[{label:'Bewerbungen',data:mData,backgroundColor:'#0b5fff'}]},
        options:{responsive:true,maintainAspectRatio:false}
      })

      const statusCounts = DEFAULT_STATUSES.map(s=>apps.filter(a=>a.status===s).length)
      if(chartStatus) chartStatus.destroy()
      chartStatus = new Chart(document.getElementById('chartStatus'),{
        type:'pie',
        data:{
          labels:DEFAULT_STATUSES,
          datasets:[{data:statusCounts,backgroundColor:['#adb5bd','#0b5fff','#6c757d','#fd7e14','#198754','#dc3545']}]
        },
        options:{responsive:true,maintainAspectRatio:false}
      })
    }
  }

  function renderDashboard(){
    try {
      // Response Rate
      const applied = apps.filter(a=>a.status!=='Entwurf').length
      const withResponse = apps.filter(a=>['Eingangsbestätigung','Interview','Angebot','Absage'].includes(a.status)).length
      const responseRate = applied > 0 ? Math.round(withResponse/applied*100) : 0
      document.getElementById('responseRate').textContent = responseRate+'%'
      document.getElementById('responseRateDesc').textContent = `${withResponse}/${applied} Bewerbungen`

      // Avg response time
      let totalDays = 0, count = 0
      apps.forEach(a=>{
        if(a.date && a.status!=='Entwurf' && a.status!=='Beworben'){
          const d1 = new Date(a.date), d2 = new Date()
          totalDays += Math.floor((d2-d1)/(1000*60*60*24))
          count++
        }
      })
      const avgTime = count > 0 ? Math.round(totalDays/count) : '-'
      document.getElementById('avgResponseTime').textContent = avgTime

      // Success rate
      const offers = apps.filter(a=>a.status==='Angebot').length
      const successRate = apps.length > 0 ? Math.round(offers/apps.length*100) : 0
      document.getElementById('successRate').textContent = successRate+'%'

      // This week
      const weekAgo = new Date(Date.now()-7*24*60*60*1000)
      const weekCount = apps.filter(a=>a.date && new Date(a.date)>weekAgo).length
      document.getElementById('weekCount').textContent = weekCount

      // Funnel
      const funnel = [
        {label:'Bewerbungen gesendet',count:applied,key:'applied'},
        {label:'Response erhalten',count:withResponse,key:'response'},
        {label:'Interviews',count:apps.filter(a=>a.status==='Interview').length,key:'interview'},
        {label:'Angebote',count:offers,key:'offer'}
      ]
      const funnelEl = document.getElementById('funnelChart')
      if(funnelEl) {
        funnelEl.innerHTML = funnel.map(f=>`<div class="funnel-bar ${f.key}">${f.label}: ${f.count}</div>`).join('')
      }

      // Trend chart - weekly
      const weeks = {}
      const today = new Date()
      for(let i=11;i>=0;i--){
        const d = new Date(today.getFullYear(),today.getMonth(),today.getDate()-i*7)
        const key = `W${d.getWeek()}`
        weeks[key] = 0
      }
      apps.forEach(a=>{
        if(a.date){
          const d = new Date(a.date)
          const key = `W${d.getWeek()}`
          if(key in weeks) weeks[key]++
        }
      })
      if(typeof Chart !== 'undefined'){
        const trendCanvas = document.getElementById('chartTrend')
        if(trendCanvas){
          if(chartTrend) chartTrend.destroy()
          chartTrend = new Chart(trendCanvas,{
            type:'line',
            data:{
              labels:Object.keys(weeks),
              datasets:[{label:'Bewerbungen pro Woche',data:Object.values(weeks),borderColor:'#0b5fff',backgroundColor:'rgba(11,95,255,0.1)',tension:0.3,fill:true}]
            },
            options:{
              responsive:true,
              maintainAspectRatio:false,
              plugins:{legend:{display:true,position:'top'}},
              scales:{
                y:{beginAtZero:true,max:Math.max(...Object.values(weeks),5),ticks:{stepSize:1}},
                x:{display:true}
              }
            }
          })
        }
      }
    } catch(e) {
      console.error('Dashboard rendering error:', e)
    }
  }

  function renderReminders(){
    remindersEl.innerHTML = ''
    const now = new Date()
    const rems = apps.filter(a=>a.reminder)
    rems.sort((a,b)=>(a.reminder||'').localeCompare(b.reminder||''))
    rems.forEach(a=>{
      const li = document.createElement('li')
      const d = new Date(a.reminder)
      li.textContent = `${d.toLocaleString()} — ${a.company} ${a.position?'- '+a.position:''}`
      if(d<now) li.classList.add('overdue')
      li.addEventListener('click',()=>openModal(a))
      remindersEl.appendChild(li)
    })
  }

  function renderAll(){
    renderFilters()
    renderList()
    renderDashboard()
  }

  // View tabs
  document.querySelectorAll('.tab-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'))
      document.querySelectorAll('.view-section').forEach(s=>s.classList.remove('active'))
      btn.classList.add('active')
      const view = btn.dataset.view+'-view'
      document.getElementById(view).classList.add('active')
      if(btn.dataset.view === 'dashboard') renderDashboard()
    })
  })

  // Form handling
  form.addEventListener('submit', async (ev)=>{
    ev.preventDefault()
    const fd = new FormData(form)
    const obj = {
      id: editingId || id(),
      company: fd.get('company'),
      position: fd.get('position'),
      date: fd.get('date'),
      status: fd.get('status'),
      priority: fd.get('priority'),
      tags: (fd.get('tags')||'').split(',').map(s=>s.trim()).filter(Boolean),
      notes: fd.get('notes'),
      reminder: fd.get('reminder')||null,
      attachments: [],
      history: [],
      notes_interviews: [],
      notes_feedback: []
    }

    const files = form.files.files
    for(const f of files){
      const b64 = await readFileAsDataURL(f)
      obj.attachments.push({name:f.name,type:f.type,data:b64})
    }

    if(editingId){
      const idx = apps.findIndex(a=>a.id===editingId)
      if(idx!==-1){
        obj.attachments = apps[idx].attachments.concat(obj.attachments)
        obj.history = apps[idx].history||[]
        obj.notes_interviews = apps[idx].notes_interviews||[]
        obj.notes_feedback = apps[idx].notes_feedback||[]
        pushHistory(obj, `Bearbeitet`)
        apps[idx] = obj
      }
    } else {
      pushHistory(obj, 'Bewerbung erstellt')
      apps.push(obj)
    }
    save()
    closeModal()
    renderAll()
  })

  // Import handling
  importBtn.addEventListener('click', openImportModal)
  cancelImport.addEventListener('click', closeImportModal)

  importFile.addEventListener('change', async (e) => {
    const file = e.target.files[0]
    if (!file) return

    try {
      const { text, boldTexts } = await extractTextFromDocx(file)
      importedData = parseImportedData(text, boldTexts)
      
      // Show preview
      document.getElementById('previewCompany').textContent = importedData.company || '(nicht gefunden)'
      document.getElementById('previewPosition').textContent = importedData.position || '(nicht gefunden)'
      document.getElementById('previewDate').textContent = importedData.date || '(nicht gefunden)'
      document.getElementById('previewNotes').textContent = importedData.notes.substring(0, 100) + '...'
      
      importPreview.classList.remove('hidden')
      confirmImport.classList.remove('hidden')
    } catch (err) {
      alert('Fehler beim Lesen der Datei: ' + err.message)
      console.error(err)
    }
  })

  confirmImport.addEventListener('click', () => {
    if (!importedData) return
    
    const obj = {
      id: id(),
      company: importedData.company,
      position: importedData.position,
      date: importedData.date,
      status: 'Entwurf',
      priority: '',
      tags: [],
      notes: importedData.notes,
      reminder: null,
      attachments: [],
      history: [],
      notes_interviews: [],
      notes_feedback: []
    }
    
    pushHistory(obj, 'Aus Word importiert')
    apps.push(obj)
    save()
    closeImportModal()
    renderAll()
    alert('✅ Bewerbung erfolgreich importiert!')
  })

  // Notes modal
  notesBtn.addEventListener('click',()=>openNotesModal())
  closeNotesBtn.addEventListener('click',()=>closeNotesModal())

  notesTabBtns.forEach(btn=>{
    btn.addEventListener('click',()=>{
      notesTabBtns.forEach(b=>b.classList.remove('active'))
      document.querySelectorAll('.notes-tab').forEach(t=>t.classList.remove('active'))
      btn.classList.add('active')
      document.getElementById(btn.dataset.notesTab+'-tab').classList.add('active')
    })
  })

  templateBtns.forEach(btn=>{
    btn.addEventListener('click',()=>{
      const template = btn.dataset.template
      if(template==='interview'){
        interviewText.value = '- Gesprächspartner:\n- Unternehmensgröße/Struktur:\n- Wichtige Punkte vom Gespräch:\n- Fragen, die mir gestellt wurden:\n- Meine Fragen an das Unternehmen:\n- Nächste Schritte:'
      } else if(template==='feedback'){
        feedbackText.value = 'Schreib dein Feedback und Learnings auf...'
      } else if(template==='followup'){
        interviewText.value = '- Wann Nachfassen?\n- Was nicht geklärt?\n- Follow-up Punkte:\n- Betreff Mail:'
      }
      setTimeout(()=>document.querySelector('[data-notes-tab="interview"]').click(),50)
    })
  })

  saveInterviewBtn.addEventListener('click',()=>{
    const app = apps.find(a=>a.id===editingId)
    if(!app || !interviewText.value.trim()) return
    app.notes_interviews = app.notes_interviews || []
    app.notes_interviews.push({date:new Date().toISOString(), content:interviewText.value})
    save()
    renderNotesUI()
    interviewText.value = ''
  })

  saveFeedbackBtn.addEventListener('click',()=>{
    const app = apps.find(a=>a.id===editingId)
    if(!app || !feedbackText.value.trim()) return
    app.notes_feedback = app.notes_feedback || []
    app.notes_feedback.push({date:new Date().toISOString(), type:feedbackType.value, content:feedbackText.value})
    save()
    renderNotesUI()
    feedbackText.value = ''
  })

  cancelBtn.addEventListener('click',()=>closeModal())
  newBtn.addEventListener('click',()=>openModal(null))
  searchEl.addEventListener('input',()=>renderList())
  filterStatus.addEventListener('change',()=>renderList())
  filterPriority.addEventListener('change',()=>renderList())
  filterTag.addEventListener('input',()=>renderList())
  sortBy.addEventListener('change',()=>renderList())

  function readFileAsDataURL(file){
    return new Promise((res,rej)=>{
      const r = new FileReader()
      r.onload = ()=>res(r.result)
      r.onerror = rej
      r.readAsDataURL(file)
    })
  }

  window.openAttachment = function(data){
    const w = window.open('')
    w.document.write(`<iframe style="width:100%;height:100vh;border:0" src="${data}"></iframe>`)
  }

  // Week helper
  Date.prototype.getWeek = function(){
    const d = new Date(Date.UTC(this.getFullYear(),this.getMonth(),this.getDate()))
    const dayNum = d.getUTCDay()+1
    d.setUTCDate(d.getUTCDate()+4-dayNum)
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1))
    return Math.ceil((((d-yearStart)/86400000)+1)/7)
  }

  load()
  renderAll()

  window._apps = apps
})()
