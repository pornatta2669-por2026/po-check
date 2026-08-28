require('dotenv').config()
const express = require('express')
const cors    = require('cors')
const { google } = require('googleapis')
const path   = require('path')
const crypto = require('crypto')
const multer = require('multer')

const app = express()
app.use(cors())
app.use(express.json({ limit: '10mb' }))
app.use(express.static(path.join(__dirname)))

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } })

const SHEET_ID   = process.env.SPREADSHEET_ID
const JWT_SECRET = process.env.JWT_SECRET || 'po-check-secret-2024'
const CONTRACT_FOLDER_ID = '1a5qvwCqFk2qcFijdLabjz-F9holEVSpu' // Drive: CT_สัญญาผ่อนคอม
const SLIP_FOLDER_ID     = '1h6t8oOD8UyehTPQRytFMrvbIGDtW-wjH' // Drive: CT_สลิปผ่อนคอม

const USERS = [
  { name: process.env.USER1_NAME || 'ปอ',       email: process.env.USER1_EMAIL || 'po@co.com',                password: process.env.USER1_PASS || 'po1234' },
  { name: process.env.USER2_NAME || 'ผู้ใช้ 2',  email: process.env.USER2_EMAIL || 'u2@co.com',               password: process.env.USER2_PASS || 'user1234' },
  { name: process.env.USER3_NAME || 'ผู้ใช้ 3',  email: process.env.USER3_EMAIL || 'u3@co.com',               password: process.env.USER3_PASS || 'user1234' },
]

// Parse dates in D/M/YYYY, DD/MM/YYYY, or YYYY-MM-DD format
function parseDate(str){
  if(!str) return null
  str=String(str).trim()
  // YYYY-MM-DD
  if(/^\d{4}-\d{1,2}-\d{1,2}/.test(str)) return new Date(str)
  // D/M/YYYY or DD/MM/YYYY (Thai format)
  const m=str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
  if(m) return new Date(parseInt(m[3]),parseInt(m[2])-1,parseInt(m[1]))
  return new Date(str)
}

function b64(s){ return Buffer.from(s).toString('base64url') }
function signToken(p){
  const h=b64(JSON.stringify({alg:'HS256',typ:'JWT'}))
  const b=b64(JSON.stringify({...p,exp:Date.now()+7*24*3600*1000}))
  const s=crypto.createHmac('sha256',JWT_SECRET).update(`${h}.${b}`).digest('base64url')
  return `${h}.${b}.${s}`
}
function verifyToken(token){
  try{
    const [h,b,s]=token.split('.')
    if(crypto.createHmac('sha256',JWT_SECRET).update(`${h}.${b}`).digest('base64url')!==s) return null
    const p=JSON.parse(Buffer.from(b,'base64url').toString())
    return p.exp<Date.now()?null:p
  }catch{return null}
}
function auth(req,res,next){
  const t=req.headers.authorization?.replace('Bearer ','')
  if(!t) return res.status(401).json({ok:false,error:'กรุณาเข้าสู่ระบบ'})
  const p=verifyToken(t)
  if(!p) return res.status(401).json({ok:false,error:'Session หมดอายุ'})
  req.user=p;next()
}

// Normalize any month/date input to YYYY-MM
function normalizeMonth(val){
  if(!val) return ''
  val=String(val).trim()
  // Format: 01-31/03/2026 or 1-31/03/2026
  const m1=val.match(/\d{1,2}-\d{1,2}\/(\d{2})\/?(\d{4})/)
  if(m1) return m1[2]+'-'+m1[1].padStart(2,'0')
  // Format: dd/mm/yyyy
  const m2=val.match(/^\d{1,2}\/(\d{2})\/(\d{4})$/)
  if(m2) return m2[2]+'-'+m2[1].padStart(2,'0')
  // Format: mm/yyyy
  const m3=val.match(/^(\d{1,2})\/(\d{4})$/)
  if(m3) return m3[2]+'-'+m3[1].padStart(2,'0')
  // Format: YYYY-MM already
  const m4=val.match(/^(\d{4})-(\d{1,2})/)
  if(m4) return m4[1]+'-'+m4[2].padStart(2,'0')
  return val
}

// Computer installment plan: price>30,000 requires a down payment of (price-30,000) up front,
// and only the remaining 30,000 is split across the installments. Otherwise split the full price.
function computeInstallmentPlan(price,installments){
  const down = price>30000 ? round2(price-30000) : 0
  const base = round2(price-down)
  const monthly = round2(base/installments)
  return {down,monthly,base}
}
function round2(n){ return Math.round(n*100)/100 }
function addMonth(ym,n){
  let [y,m]=ym.split('-').map(Number)
  m+=n
  while(m>12){m-=12;y++}
  while(m<1){m+=12;y--}
  return y+'-'+String(m).padStart(2,'0')
}
function buildSchedule(recordId,price,installments,startMonth){
  const {down,monthly,base}=computeInstallmentPlan(price,installments)
  const rows=[]
  if(down>0) rows.push([recordId,0,startMonth,down,'FALSE','','',''])
  let allocated=0
  for(let i=1;i<=installments;i++){
    const amt = i===installments ? round2(base-allocated) : monthly
    allocated=round2(allocated+amt)
    rows.push([recordId,i,addMonth(startMonth,i-1),amt,'FALSE','','',''])
  }
  return rows
}

function ga(){
  return new google.auth.GoogleAuth({
    credentials:JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes:['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
  })
}
function drive(){ return google.drive({version:'v3',auth:ga()}) }
async function driveUpload(name,parentId,buffer,mimeType){
  const r=await drive().files.create({
    requestBody:{name,parents:[parentId]},
    media:{mimeType:mimeType||'application/octet-stream',body:require('stream').Readable.from(buffer)},
    fields:'id,webViewLink'
  })
  return {id:r.data.id,url:r.data.webViewLink}
}
async function driveCreateFolder(name,parentId){
  const r=await drive().files.create({requestBody:{name,mimeType:'application/vnd.google-apps.folder',parents:[parentId]},fields:'id'})
  return r.data.id
}
async function readSheet(range){
  const r=await google.sheets({version:'v4',auth:ga()}).spreadsheets.values.get({spreadsheetId:SHEET_ID,range})
  return r.data.values||[]
}
async function appendRow(sheet,row){
  await google.sheets({version:'v4',auth:ga()}).spreadsheets.values.append({
    spreadsheetId:SHEET_ID,range:`${sheet}!A1`,
    valueInputOption:'USER_ENTERED',requestBody:{values:[row]}
  })
}
async function updateRow(sheet,idx,row){
  await google.sheets({version:'v4',auth:ga()}).spreadsheets.values.update({
    spreadsheetId:SHEET_ID,range:`${sheet}!A${idx}`,
    valueInputOption:'USER_ENTERED',requestBody:{values:[row]}
  })
}
async function deleteRow(sheet,idx){
  const meta=await google.sheets({version:'v4',auth:ga()}).spreadsheets.get({spreadsheetId:SHEET_ID})
  const sh=meta.data.sheets.find(s=>s.properties.title===sheet)
  if(!sh) throw new Error('Sheet not found: '+sheet)
  await google.sheets({version:'v4',auth:ga()}).spreadsheets.batchUpdate({
    spreadsheetId:SHEET_ID,
    requestBody:{requests:[{deleteDimension:{range:{sheetId:sh.properties.sheetId,dimension:'ROWS',startIndex:idx-1,endIndex:idx}}}]}
  })
}
async function log(user,action,detail){
  try{await appendRow('LOGS',[new Date().toLocaleString('th-TH'),user.name,user.email,action,detail])}
  catch(e){console.error('Log:',e.message)}
}

// AUTH
app.post('/api/auth/login',async(req,res)=>{
  const{email,password}=req.body
  if(!email||!password) return res.status(400).json({ok:false,error:'กรุณากรอกข้อมูล'})
  const byEmail=USERS.find(u=>u.email.toLowerCase()===email.toLowerCase())
  if(!byEmail) return res.status(401).json({ok:false,error:'ไม่พบ Email นี้ในระบบ กรุณาตรวจสอบอีเมลอีกครั้ง'})
  if(byEmail.password!==password) return res.status(401).json({ok:false,error:'รหัสผ่านไม่ถูกต้อง กรุณาตรวจสอบรหัสผ่านอีกครั้ง'})
  const u=byEmail
  await log({name:u.name,email:u.email},'LOGIN','เข้าสู่ระบบ')
  res.json({ok:true,token:signToken({name:u.name,email:u.email}),user:{name:u.name,email:u.email}})
})
app.get('/api/auth/me',auth,(req,res)=>res.json({ok:true,user:req.user}))

// EMP_PO / EMP_PO_ITEMS — new Employee/Department/PO module. Separate sheets from PO_MASTER/PO_EMPLOYEES; never touch those.
app.get('/api/emp-po',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('EMP_PO!A2:J')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.get('/api/emp-po-items',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('EMP_PO_ITEMS!A2:G')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.get('/api/emp-employees',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('EMPLOYEES!A2:G')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.get('/api/emp-departments',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('DEPARTMENTS!A2:E')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/emp-po/:poNo/month',auth,async(req,res)=>{
  try{
    const{month}=req.body
    if(!month) return res.status(400).json({ok:false,error:'กรุณาระบุเดือน'})
    const rows=await readSheet('EMP_PO!A2:J')
    const i=rows.findIndex(r=>r[0]===req.params.poNo)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบ PO'})
    const row=rows[i]
    const expiryOrEnd=row[5]||row[4]
    let approved
    if((row[9]||'').trim()){
      approved=new Set(row[9].split(',').map(m=>m.trim()).filter(Boolean))
    }else if(row[6]==='Approved'){
      // migrate implicit whole-PO "Approved" into an explicit per-month set spanning the PO's own range
      approved=new Set()
      if(row[3]){
        let[y,m]=row[3].slice(0,7).split('-').map(Number)
        const endM=expiryOrEnd?expiryOrEnd.slice(0,7):row[3].slice(0,7)
        while(y+'-'+String(m).padStart(2,'0')<=endM){approved.add(y+'-'+String(m).padStart(2,'0'));m++;if(m>12){m=1;y++}}
      }
    }else{
      approved=new Set()
    }
    if(approved.has(month)) approved.delete(month); else approved.add(month)
    const updated=[...row]
    updated[9]=Array.from(approved).sort().join(',')
    await updateRow('EMP_PO',i+2,updated.slice(0,10))
    await log(req.user,'EMP_PO_MONTH',`${approved.has(month)?'อนุมัติ':'ยกเลิกอนุมัติ'} PO ${req.params.poNo} เดือน ${month}`)
    res.json({ok:true,approvedMonths:updated[9]})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/emp-departments',auth,async(req,res)=>{
  try{
    const{name,head}=req.body
    if(!name) return res.status(400).json({ok:false,error:'กรุณาระบุชื่อแผนก'})
    const id='DEPT'+Date.now()
    await appendRow('DEPARTMENTS',[id,name,head||'',req.user.name,new Date().toLocaleString('th-TH')])
    await log(req.user,'CREATE_EMP_DEPT',`สร้างแผนก ${name}`)
    res.json({ok:true,id})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/emp-departments/:id',auth,async(req,res)=>{
  try{
    const{name,head}=req.body
    if(!name) return res.status(400).json({ok:false,error:'กรุณาระบุชื่อแผนก'})
    const rows=await readSheet('DEPARTMENTS!A2:E')
    const i=rows.findIndex(r=>r[0]===req.params.id)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบแผนก'})
    const updated=[...rows[i]];updated[1]=name;updated[2]=head||''
    await updateRow('DEPARTMENTS',i+2,updated.slice(0,5))
    await log(req.user,'UPDATE_EMP_DEPT',`แก้ไขแผนก ${req.params.id} เป็น ${name}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/emp-employees/bulk-department',auth,async(req,res)=>{
  try{
    const{rows}=req.body
    if(!rows||!rows.length) return res.status(400).json({ok:false,error:'ไม่มีข้อมูล'})
    const str=v=>(v===null||v===undefined)?'':String(v).trim()
    const[emps,depts]=await Promise.all([readSheet('EMPLOYEES!A2:G'),readSheet('DEPARTMENTS!A2:E')])
    const now=new Date().toLocaleString('th-TH')
    let updated=0,deptCreated=0,dCounter=0;const notFoundIds=[];const updateData=[];const newDeptRows=[]
    for(const r of rows){
      const id=str(r['Employee ID']??r['EmployeeID']??r['employee id']??'')
      if(!id) continue
      const idx=emps.findIndex(e=>e[0]===id)
      if(idx===-1){notFoundIds.push(id);continue}
      const deptName=str(r['แผนก']??r['Department']??r['department']??'')
      if(!deptName) continue
      let dept=depts.find(d=>str(d[1]).toLowerCase()===deptName.toLowerCase())
      if(!dept){
        dept=['DEPT'+Date.now()+(dCounter++),deptName,'',req.user.name,now]
        depts.push(dept)
        newDeptRows.push(dept)
        deptCreated++
      }
      const row=[...emps[idx]];row[2]=dept[0]
      updateData.push({range:`EMPLOYEES!A${idx+2}`,values:[row.slice(0,7)]})
      emps[idx]=row
      updated++
    }
    const sheets=google.sheets({version:'v4',auth:ga()})
    if(newDeptRows.length) await sheets.spreadsheets.values.append({spreadsheetId:SHEET_ID,range:'DEPARTMENTS!A1',valueInputOption:'USER_ENTERED',requestBody:{values:newDeptRows}})
    if(updateData.length) await sheets.spreadsheets.values.batchUpdate({spreadsheetId:SHEET_ID,requestBody:{valueInputOption:'USER_ENTERED',data:updateData}})
    await log(req.user,'EMP_BULK_DEPT',`อัปเดตแผนกพนักงาน: สำเร็จ ${updated} ไม่พบพนักงาน ${notFoundIds.length} สร้างแผนกใหม่ ${deptCreated}`)
    res.json({ok:true,updated,notFoundIds,deptCreated})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/emp-employees/:id/department',auth,async(req,res)=>{
  try{
    const{departmentId}=req.body
    const rows=await readSheet('EMPLOYEES!A2:G')
    const i=rows.findIndex(r=>r[0]===req.params.id)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบพนักงาน'})
    const updated=[...rows[i]];updated[2]=departmentId||''
    await updateRow('EMPLOYEES',i+2,updated.slice(0,7))
    await log(req.user,'EMP_SET_DEPT',`ย้าย ${req.params.id} เข้าแผนก ${departmentId}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/emp-employees/:id/resign',auth,async(req,res)=>{
  try{
    const{month}=req.body
    if(!month) return res.status(400).json({ok:false,error:'กรุณาระบุเดือนที่ลาออก'})
    const rows=await readSheet('EMPLOYEES!A2:G')
    const i=rows.findIndex(r=>r[0]===req.params.id)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบพนักงาน'})
    const updated=[...rows[i]];updated[3]='Resigned';updated[4]=month
    await updateRow('EMPLOYEES',i+2,updated.slice(0,7))
    await log(req.user,'EMP_RESIGN',`บันทึกลาออก ${req.params.id} มีผล ${month}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/emp-employees/bulk',auth,async(req,res)=>{
  try{
    const{rows}=req.body
    if(!rows||!rows.length) return res.status(400).json({ok:false,error:'ไม่มีข้อมูล'})
    const str=v=>(v===null||v===undefined)?'':String(v).trim()
    const[existing,depts]=await Promise.all([readSheet('EMPLOYEES!A2:A'),readSheet('DEPARTMENTS!A2:E')])
    const existingIds=new Set(existing.map(r=>str(r[0])).filter(Boolean))
    const now=new Date().toLocaleString('th-TH')
    const newRows=[];const skippedIds=[]
    for(const r of rows){
      const id=str(r['Employee ID']??r['EmployeeID']??r['employee id']??'')
      if(!id) continue
      if(existingIds.has(id)){skippedIds.push(id);continue}
      const name=str(r['ชื่อ-นามสกุล']??r['Name']??r['name']??'')
      const deptName=str(r['แผนก']??r['Department']??r['department']??'')
      const dept=depts.find(d=>str(d[1]).toLowerCase()===deptName.toLowerCase())
      const status=str(r['สถานะ']??r['Status']??'Active')||'Active'
      newRows.push([id,name,dept?dept[0]:'',status,'',req.user.name,now])
      existingIds.add(id)
    }
    if(newRows.length){
      const sheets=google.sheets({version:'v4',auth:ga()})
      await sheets.spreadsheets.values.append({spreadsheetId:SHEET_ID,range:'EMPLOYEES!A1',valueInputOption:'USER_ENTERED',requestBody:{values:newRows}})
    }
    await log(req.user,'EMP_BULK_IMPORT',`นำเข้าพนักงาน: เพิ่ม ${newRows.length} ข้าม (ซ้ำ) ${skippedIds.length}`)
    res.json({ok:true,added:newRows.length,skipped:skippedIds.length,skippedIds})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/emp-po',auth,async(req,res)=>{
  try{
    const{poNo,client,departmentId,startDate,endDate,expireDate,items}=req.body
    if(!poNo) return res.status(400).json({ok:false,error:'กรุณากรอกเลขที่ PO'})
    if(!items||!items.length||!items.some(it=>it.employeeId&&it.start&&it.end)) return res.status(400).json({ok:false,error:'กรุณาเพิ่มพนักงานอย่างน้อย 1 คน พร้อมวันที่ให้ครบ'})
    const existing=await readSheet('EMP_PO!A2:A')
    if(existing.some(r=>r[0]===poNo)) return res.status(400).json({ok:false,error:'มีเลขที่ PO นี้อยู่แล้ว'})
    const now=new Date().toLocaleString('th-TH')
    await appendRow('EMP_PO',[poNo,client||'',departmentId||'',startDate||'',endDate||'',expireDate||'','Waiting',req.user.name,now])
    const emps=await readSheet('EMPLOYEES!A2:G')
    const empName=id=>{const e=emps.find(x=>x[0]===id);return e?e[1]:''}
    const validItems=items.filter(it=>it.employeeId&&it.start&&it.end)
    const itemRows=validItems.map(it=>[poNo,it.employeeId,empName(it.employeeId),it.start,it.end,req.user.name,now])
    if(itemRows.length){
      const sheets=google.sheets({version:'v4',auth:ga()})
      await sheets.spreadsheets.values.append({spreadsheetId:SHEET_ID,range:'EMP_PO_ITEMS!A1',valueInputOption:'USER_ENTERED',requestBody:{values:itemRows}})
    }
    await log(req.user,'CREATE_EMP_PO',`สร้าง PO ${poNo} (${itemRows.length} คน)`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.put('/api/emp-po/:poNo',auth,async(req,res)=>{
  try{
    const{client,departmentId,startDate,endDate,expireDate,items}=req.body
    if(!items||!items.length||!items.some(it=>it.employeeId&&it.start&&it.end)) return res.status(400).json({ok:false,error:'กรุณาเพิ่มพนักงานอย่างน้อย 1 คน พร้อมวันที่ให้ครบ'})
    const rows=await readSheet('EMP_PO!A2:J')
    const i=rows.findIndex(r=>r[0]===req.params.poNo)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบ PO'})
    const orig=rows[i]
    await updateRow('EMP_PO',i+2,[req.params.poNo,client||'',departmentId||'',startDate||'',endDate||'',expireDate||'',orig[6],orig[7],orig[8],orig[9]||''])
    const sheets=google.sheets({version:'v4',auth:ga()})
    const itemRows=await readSheet('EMP_PO_ITEMS!A2:G')
    const delIdx=[]
    itemRows.forEach((r,idx)=>{if(r[0]===req.params.poNo) delIdx.push(idx+2)})
    if(delIdx.length){
      const meta=await sheets.spreadsheets.get({spreadsheetId:SHEET_ID})
      const sh=meta.data.sheets.find(s=>s.properties.title==='EMP_PO_ITEMS')
      const requests=delIdx.sort((a,b)=>b-a).map(idx=>({deleteDimension:{range:{sheetId:sh.properties.sheetId,dimension:'ROWS',startIndex:idx-1,endIndex:idx}}}))
      await sheets.spreadsheets.batchUpdate({spreadsheetId:SHEET_ID,requestBody:{requests}})
    }
    const emps=await readSheet('EMPLOYEES!A2:G')
    const empName=id=>{const e=emps.find(x=>x[0]===id);return e?e[1]:''}
    const now=new Date().toLocaleString('th-TH')
    const validItems=items.filter(it=>it.employeeId&&it.start&&it.end)
    const newItemRows=validItems.map(it=>[req.params.poNo,it.employeeId,empName(it.employeeId),it.start,it.end,req.user.name,now])
    if(newItemRows.length) await sheets.spreadsheets.values.append({spreadsheetId:SHEET_ID,range:'EMP_PO_ITEMS!A1',valueInputOption:'USER_ENTERED',requestBody:{values:newItemRows}})
    await log(req.user,'UPDATE_EMP_PO',`แก้ไข PO ${req.params.poNo} (${newItemRows.length} คน)`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})

// COMPUTER_INSTALLMENT / COMPUTER_INSTALLMENT_SCHEDULE — ผ่อนคอมพิวเตอร์. Separate new sheets; Drive folders shared with the service account.
app.get('/api/computer',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('COMPUTER_INSTALLMENT!A2:N')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.get('/api/computer-schedule',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('COMPUTER_INSTALLMENT_SCHEDULE!A2:H')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/computer',auth,async(req,res)=>{
  try{
    const{employeeId,employeeName,price,installments,startMonth}=req.body
    const priceNum=Number(price),instNum=parseInt(installments)||10
    if(!employeeId||!employeeName||!priceNum||priceNum<=0||!startMonth) return res.status(400).json({ok:false,error:'กรุณากรอกข้อมูลให้ครบ (พนักงาน, ราคาเครื่อง, เดือนเริ่มผ่อน)'})
    const existing=await readSheet('COMPUTER_INSTALLMENT!A2:N')
    const machineNo=existing.filter(r=>r[1]===employeeId).length+1
    const {down,monthly}=computeInstallmentPlan(priceNum,instNum)
    const folderName=`${employeeId}_${employeeName}${machineNo>1?' เครื่องที่'+machineNo:''}`
    const slipFolderId=await driveCreateFolder(folderName,SLIP_FOLDER_ID)
    const id='COMP'+Date.now()
    const now=new Date().toLocaleString('th-TH')
    await appendRow('COMPUTER_INSTALLMENT',[id,employeeId,employeeName,machineNo,priceNum,instNum,down,monthly,startMonth,'','',slipFolderId,req.user.name,now])
    const scheduleRows=buildSchedule(id,priceNum,instNum,startMonth)
    await google.sheets({version:'v4',auth:ga()}).spreadsheets.values.append({spreadsheetId:SHEET_ID,range:'COMPUTER_INSTALLMENT_SCHEDULE!A1',valueInputOption:'USER_ENTERED',requestBody:{values:scheduleRows}})
    await log(req.user,'CREATE_COMPUTER',`เพิ่มรายการผ่อนคอม ${employeeId} ${employeeName} เครื่องที่ ${machineNo} ราคา ${priceNum}`)
    res.json({ok:true,id})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.put('/api/computer/:id',auth,async(req,res)=>{
  try{
    const{price,installments,startMonth}=req.body
    const priceNum=Number(price),instNum=parseInt(installments)||10
    if(!priceNum||priceNum<=0||!startMonth) return res.status(400).json({ok:false,error:'กรุณากรอกข้อมูลให้ครบ'})
    const rows=await readSheet('COMPUTER_INSTALLMENT!A2:N')
    const i=rows.findIndex(r=>r[0]===req.params.id)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    const schedule=await readSheet('COMPUTER_INSTALLMENT_SCHEDULE!A2:H')
    const mine=schedule.map((r,idx)=>({r,idx})).filter(x=>x.r[0]===req.params.id)
    if(mine.some(x=>x.r[4]==='TRUE')) return res.status(400).json({ok:false,error:'มีงวดที่จ่ายแล้ว ไม่สามารถแก้ไขราคา/จำนวนงวดได้ กรุณาลบแล้วสร้างใหม่หากจำเป็น'})
    const orig=rows[i]
    const {down,monthly}=computeInstallmentPlan(priceNum,instNum)
    const updated=[...orig];updated[4]=priceNum;updated[5]=instNum;updated[6]=down;updated[7]=monthly;updated[8]=startMonth
    await updateRow('COMPUTER_INSTALLMENT',i+2,updated.slice(0,14))
    if(mine.length){
      const sheets=google.sheets({version:'v4',auth:ga()})
      const meta=await sheets.spreadsheets.get({spreadsheetId:SHEET_ID})
      const sh=meta.data.sheets.find(s=>s.properties.title==='COMPUTER_INSTALLMENT_SCHEDULE')
      const requests=mine.map(x=>x.idx+2).sort((a,b)=>b-a).map(rowNum=>({deleteDimension:{range:{sheetId:sh.properties.sheetId,dimension:'ROWS',startIndex:rowNum-1,endIndex:rowNum}}}))
      await sheets.spreadsheets.batchUpdate({spreadsheetId:SHEET_ID,requestBody:{requests}})
    }
    const scheduleRows=buildSchedule(req.params.id,priceNum,instNum,startMonth)
    await google.sheets({version:'v4',auth:ga()}).spreadsheets.values.append({spreadsheetId:SHEET_ID,range:'COMPUTER_INSTALLMENT_SCHEDULE!A1',valueInputOption:'USER_ENTERED',requestBody:{values:scheduleRows}})
    await log(req.user,'UPDATE_COMPUTER',`แก้ไขรายการผ่อนคอม ${req.params.id}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.delete('/api/computer/:id',auth,async(req,res)=>{
  try{
    const rows=await readSheet('COMPUTER_INSTALLMENT!A2:N')
    const i=rows.findIndex(r=>r[0]===req.params.id)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    await deleteRow('COMPUTER_INSTALLMENT',i+2)
    const schedule=await readSheet('COMPUTER_INSTALLMENT_SCHEDULE!A2:H')
    const delIdx=schedule.map((r,idx)=>({r,idx})).filter(x=>x.r[0]===req.params.id).map(x=>x.idx+2)
    if(delIdx.length){
      const sheets=google.sheets({version:'v4',auth:ga()})
      const meta=await sheets.spreadsheets.get({spreadsheetId:SHEET_ID})
      const sh=meta.data.sheets.find(s=>s.properties.title==='COMPUTER_INSTALLMENT_SCHEDULE')
      const requests=delIdx.sort((a,b)=>b-a).map(rowNum=>({deleteDimension:{range:{sheetId:sh.properties.sheetId,dimension:'ROWS',startIndex:rowNum-1,endIndex:rowNum}}}))
      await sheets.spreadsheets.batchUpdate({spreadsheetId:SHEET_ID,requestBody:{requests}})
    }
    await log(req.user,'DELETE_COMPUTER',`ลบรายการผ่อนคอม ${req.params.id} (${rows[i][1]} ${rows[i][2]})`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/computer/:id/contract',auth,upload.single('file'),async(req,res)=>{
  try{
    if(!req.file) return res.status(400).json({ok:false,error:'กรุณาแนบไฟล์'})
    const rows=await readSheet('COMPUTER_INSTALLMENT!A2:N')
    const i=rows.findIndex(r=>r[0]===req.params.id)
    if(i===-1) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    const row=rows[i]
    const machineSuffix=Number(row[3])>1?' เครื่องที่'+row[3]:''
    const ext=(req.file.originalname.match(/\.[^.]+$/)||['.pdf'])[0]
    const name=`${row[1]}_สัญญาผ่อนคอม_${row[2]}${machineSuffix}${ext}`
    const {id:fileId,url}=await driveUpload(name,CONTRACT_FOLDER_ID,req.file.buffer,req.file.mimetype)
    const updated=[...row];updated[9]=fileId;updated[10]=url
    await updateRow('COMPUTER_INSTALLMENT',i+2,updated.slice(0,14))
    await log(req.user,'COMPUTER_CONTRACT',`แนบสัญญาผ่อนคอม ${req.params.id} (${row[1]} ${row[2]})`)
    res.json({ok:true,url})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/computer-schedule/:rowIdx/paid',auth,upload.single('file'),async(req,res)=>{
  try{
    const idx=parseInt(req.params.rowIdx)
    const schedule=await readSheet('COMPUTER_INSTALLMENT_SCHEDULE!A2:H')
    const row=schedule[idx-2]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบงวด'})
    const turningOn = row[4]!=='TRUE'
    const updated=[...row]
    updated[4]=turningOn?'TRUE':'FALSE'
    if(turningOn){
      updated[5]=new Date().toLocaleDateString('th-TH')
      if(req.file){
        const records=await readSheet('COMPUTER_INSTALLMENT!A2:N')
        const rec=records.find(r=>r[0]===row[0])
        if(rec&&rec[11]){
          const ext=(req.file.originalname.match(/\.[^.]+$/)||['.jpg'])[0]
          const label=Number(row[1])===0?'ดาวน์':'งวด'+String(row[1]).padStart(2,'0')
          const {id:fileId,url}=await driveUpload(label+ext,rec[11],req.file.buffer,req.file.mimetype)
          updated[6]=fileId;updated[7]=url
        }
      }
    }
    await updateRow('COMPUTER_INSTALLMENT_SCHEDULE',idx,updated.slice(0,8))
    await log(req.user,'COMPUTER_TOGGLE_PAID',`${turningOn?'ทำเครื่องหมายจ่ายแล้ว':'ยกเลิกจ่ายแล้ว'} ${row[0]} งวด ${row[1]}`)
    res.json({ok:true,paid:updated[4]})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})

// OT
app.get('/api/ot',auth,async(req,res)=>{
  try{res.json({ok:true,data:await readSheet('OT!A2:K')})}
  catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/ot',auth,async(req,res)=>{
  try{
    const{poNum,empName,poOt,docDate,otPay,otBill,otMonth,yearRemark}=req.body
    if(!poNum||!empName||!docDate||!otPay||!otBill) return res.status(400).json({ok:false,error:'กรุณากรอกข้อมูลที่บังคับ (*)'})
    const month=otMonth||docDate.slice(0,7)
    const existing=await readSheet('OT!A2:K')
    if(existing.some(r=>r[0]===poNum && r[1]===empName && r[7]===month)) return res.status(400).json({ok:false,error:'เดือนนี้ถูกเบิกไปแล้วสำหรับ PO และพนักงานคนนี้'})
    await appendRow('OT',[poNum,empName,poOt||'',docDate,otPay,otBill,Number(otBill)-Number(otPay),month,'ยังไม่เรียกเก็บ',req.user.name,yearRemark||''])
    await log(req.user,'CREATE_OT',`บันทึก OT ${empName} PO: ${poNum} เดือน: ${month}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/ot/:rowIdx/poot',auth,async(req,res)=>{
  try{
    const idx=parseInt(req.params.rowIdx)
    const rows=await readSheet('OT!A2:K');const row=rows[idx-1]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    const updated=[...row];updated[2]=req.body.poOt||''
    await updateRow('OT',idx+1,updated.slice(0,10))
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.patch('/api/ot/:rowIdx/collected',auth,async(req,res)=>{
  try{
    const idx=parseInt(req.params.rowIdx)
    const rows=await readSheet('OT!A2:K');const row=rows[idx-1]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    const newStatus=row[8]==='เรียกเก็บแล้ว'?'ยังไม่เรียกเก็บ':'เรียกเก็บแล้ว'
    const updated=[...row];updated[8]=newStatus
    await updateRow('OT',idx+1,updated.slice(0,10))
    res.json({ok:true,status:newStatus})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})

app.put('/api/ot/:rowIdx',auth,async(req,res)=>{
  try{
    const idx=parseInt(req.params.rowIdx)
    const rows=await readSheet('OT!A2:K');const row=rows[idx-1]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    const{otPay,otBill,poOt}=req.body
    const pay=otPay!==undefined?Number(otPay):Number(row[4])
    const bill=otBill!==undefined?Number(otBill):Number(row[5])
    const updated=[...row]
    if(poOt!==undefined) updated[2]=poOt
    updated[4]=pay;updated[5]=bill;updated[6]=bill-pay
    await updateRow('OT',idx+1,updated.slice(0,10))
    await log(req.user,'UPDATE_OT',`แก้ไขยอด OT ${row[1]} PO: ${row[0]}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.delete('/api/ot/:rowIdx',auth,async(req,res)=>{
  try{
    const idx=parseInt(req.params.rowIdx)
    const rows=await readSheet('OT!A2:K');const row=rows[idx-1]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    await deleteRow('OT',idx+1)
    await log(req.user,'DELETE_OT',`ลบ OT ${row[1]} PO: ${row[0]}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})

// TRACKING
// cols: PO|ประเภท|MM/YYYY|DM|StatusCT|วันที่ส่ง|Note|TrackingNo|StatusAIS|PO_OT_No|วันที่completed|MaterialDoc|CreatedBy|CreatedAt
app.get('/api/tracking',auth,async(req,res)=>{
  try{
    const rows=await readSheet('TRACKING!A2:N')
    const{dm,type,status}=req.query
    let data=rows.map((r,i)=>({
      idx:i+2,po:r[0]||'',type:r[1]||'',month:r[2]||'',dm:r[3]||'',
      statusCT:r[4]||'',sentDate:r[5]||'',note:r[6]||'',trackingNo:r[7]||'',
      statusAIS:r[8]||'',poOtNo:r[9]||'',completedDate:r[10]||'',
      materialDoc:r[11]||'',createdBy:r[12]||'',createdAt:r[13]||''
    })).filter(r=>r.po)
    if(dm) data=data.filter(r=>r.dm===dm)
    if(type) data=data.filter(r=>r.type===type)
    if(status) data=data.filter(r=>r.statusAIS===status)
    res.json({ok:true,data})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.get('/api/tracking/summary',auth,async(req,res)=>{
  try{
    const rows=await readSheet('TRACKING!A2:N')
    const data=rows.filter(r=>r[0])
    const waiting=data.filter(r=>r[8]==='Waiting for AIS to proceed.').length
    const completed=data.filter(r=>r[8]==='AIS processing completed').length
    const reject=data.filter(r=>r[8]==='Reject').length
    res.json({ok:true,total:data.length,waiting,completed,reject})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/tracking',auth,async(req,res)=>{
  try{
    const{po,type,month,dm,statusCT,sentDate,note,trackingNo,statusAIS,poOtNo,completedDate,materialDoc}=req.body
    if(!po||!trackingNo||!statusAIS) return res.status(400).json({ok:false,error:'กรุณากรอก เลข PO, Tracking No. และ Status AIS'})
    await appendRow('TRACKING',[po,type||'',month||'',dm||'',statusCT||'',sentDate||'',note||'',trackingNo,statusAIS,poOtNo||'',completedDate||'',materialDoc||'',req.user.name,new Date().toLocaleString('th-TH')])
    await log(req.user,'CREATE_TRACKING',`เพิ่ม Tracking ${trackingNo} PO: ${po}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.put('/api/tracking/:rowIdx',auth,async(req,res)=>{
  try{
    const sheetRow=parseInt(req.params.rowIdx)
    const rows=await readSheet('TRACKING!A2:N')
    const row=rows[sheetRow-2]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    const{po,type,month,dm,statusCT,sentDate,note,trackingNo,statusAIS,poOtNo,completedDate,materialDoc}=req.body
    await updateRow('TRACKING',sheetRow,[
      po||row[0],type??row[1],month??row[2],dm??row[3],statusCT??row[4],
      row[5]||(sentDate||''),note??row[6],trackingNo||row[7],statusAIS||row[8],
      poOtNo??row[9],completedDate??row[10],materialDoc??row[11],
      req.user.name,new Date().toLocaleString('th-TH')
    ])
    await log(req.user,'UPDATE_TRACKING',`แก้ไข Tracking ${trackingNo||row[7]} PO: ${po||row[0]}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.delete('/api/tracking/:rowIdx',auth,async(req,res)=>{
  try{
    const sheetRow=parseInt(req.params.rowIdx)
    const rows=await readSheet('TRACKING!A2:N')
    const row=rows[sheetRow-2]
    if(!row) return res.status(404).json({ok:false,error:'ไม่พบรายการ'})
    await deleteRow('TRACKING',sheetRow)
    await log(req.user,'DELETE_TRACKING',`ลบ Tracking ${row[7]} PO: ${row[0]}`)
    res.json({ok:true})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})
app.post('/api/tracking/bulk',auth,async(req,res)=>{
  try{
    const{rows}=req.body
    if(!rows||!rows.length) return res.status(400).json({ok:false,error:'ไม่มีข้อมูล'})
    const existing=await readSheet('TRACKING!A2:N')
    const existingSet=new Set(existing.filter(r=>r[7]).map(r=>r[7].trim()))
    const now=new Date().toLocaleString('th-TH')
    let skipped=0
    const newRows=[]
    for(const r of rows){
      const trackingNo=String(r.trackingNo||'').trim()
      if(!r.po||!trackingNo){skipped++;continue}
      if(existingSet.has(trackingNo)){skipped++;continue}
      newRows.push([
        r.po,r.type||'',r.month||'',r.dm||'',r.statusCT||'',
        r.sentDate||'',r.note||'',trackingNo,r.statusAIS||'',
        r.poOtNo||'',r.completedDate||'',r.materialDoc||'',
        req.user.name,now
      ])
      existingSet.add(trackingNo)
    }
    if(newRows.length){
      await google.sheets({version:'v4',auth:ga()}).spreadsheets.values.append({
        spreadsheetId:SHEET_ID,range:'TRACKING!A1',
        valueInputOption:'USER_ENTERED',
        requestBody:{values:newRows}
      })
    }
    await log(req.user,'TRACKING_BULK',`Import Tracking: เพิ่ม ${newRows.length} ข้าม ${skipped} รายการ`)
    res.json({ok:true,added:newRows.length,skipped})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})

// LOGS
app.get('/api/logs',auth,async(req,res)=>{
  try{
    const rows=await readSheet('LOGS!A2:E')
    res.json({ok:true,data:[...rows].reverse().slice(0,200)})
  }catch(e){res.status(500).json({ok:false,error:e.message})}
})

app.get('/api/health',(req,res)=>res.json({ok:true}))
app.get('/landing',(req,res)=>res.sendFile(path.join(__dirname,'landing.html')))
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'index.html')))

const PORT=process.env.PORT||8080
app.listen(PORT,()=>console.log(`✅ PO System on port ${PORT}`))
