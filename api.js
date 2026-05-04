const SPREADSHEET_ID = '1gWs8V-9xy7nsy54QkOrSnrrcr6DiDMcEYMINp-zEdjo';

function doGet(e) {
  const p = e.parameter;
  const action = p.action;
  let result;

  try {
    switch(action) {
      case 'book':
        result = book(p);
        break;
      case 'check':
        result = check(p);
        break;
      case 'getOpenDates':
        result = getOpenDates();
        break;
      case 'getSlots':
        result = getSlots(p);
        break;
      case 'getRawSlots':
        result = getRawSlots(p);
        break;
      case 'openDate':
        result = openDate(p);
        break;
      case 'closeDate':
        result = closeDate(p);
        break;
      case 'openAllMonth':
        result = openAllMonth(p);
        break;
      case 'closeAllMonth':
        result = closeAllMonth(p);
        break;
      case 'openSlot':
        result = openSlot(p);
        break;
      case 'closeSlot':
        result = closeSlot(p);
        break;
      case 'getBookings':
        result = getBookings(p);
        break;
      case 'confirmBooking':
        result = confirmBooking(p);
        break;
      case 'cancelBooking':
        result = cancelBooking(p);
        break;
      case 'deleteBooking':
        result = deleteBooking(p);
        break;
      case 'moveBooking':
        result = moveBooking(p);
        break;
      case 'blacklistAdd':
        result = blacklistAdd(p);
        break;
      case 'blacklistRemove':
        result = blacklistRemove(p);
        break;
      case 'getBlacklist':
        result = getBlacklist();
        break;
      case 'clientHistory':
        result = clientHistory(p);
        break;
      case 'saveCost':
        result = saveCost(p);
        break;
      case 'stats':
        result = stats(p);
        break;
      case 'crmSearch':
        result = crmSearch(p);
        break;
      case 'saveClientNote':
        result = saveClientNote(p);
        break;
      default:
        result = { ok: true, message: 'Zhosik Nails API v2' };
    }
  } catch(err) {
    result = { ok: false, error: err.toString() };
  }

  const json = JSON.stringify(result);
  const callback = p.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (name === 'Bookings') s.getRange(1,1,1,10).setValues([['id','service','serviceName','date','time','duration','clientName','clientPhone','status','createdAt']]);
    else if (name === 'OpenDates') s.getRange(1,1,1,1).setValues([['date']]);
    else if (name === 'OpenSlots') s.getRange(1,1,1,2).setValues([['date','time']]);
    else if (name === 'Blacklist') s.getRange(1,1,1,3).setValues([['phone','reason','addedAt']]);
    else if (name === 'Costs') s.getRange(1,1,1,4).setValues([['bookingId','cost','note','createdAt']]);
    else if (name === 'Clients') s.getRange(1,1,1,5).setValues([['clientPhone','clientName','note','tags','lastVisit']]);
  }
  return s;
}

function genId() { return 'b' + Date.now() + '_' + Math.random().toString(36).slice(2,8); }

function isBlacklisted(phone) {
  const data = getSheet('Blacklist').getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0]===phone) return true;
  return false;
}

function getOpenDatesArr() {
  const data = getSheet('OpenDates').getDataRange().getValues();
  const arr = [];
  for (let i=1; i<data.length; i++) if (data[i][0]) arr.push(data[i][0]);
  return arr.sort();
}

function getSlotsArr(date) {
  const data = getSheet('OpenSlots').getDataRange().getValues();
  const arr = [];
  for (let i=1; i<data.length; i++) if (data[i][0]===date) arr.push(data[i][1]);
  return arr.sort();
}

function timeToMin(t) { const x=t.split(':'); return parseInt(x[0])*60+parseInt(x[1]); }

function getBookingsForDate(date) {
  const data = getSheet('Bookings').getDataRange().getValues();
  const arr = [];
  for (let i=1; i<data.length; i++) {
    if (data[i][3]===date && data[i][8]!=='cancelled') {
      arr.push({ id:data[i][0], time:data[i][4], duration:parseInt(data[i][5])||0 });
    }
  }
  return arr;
}

function isConflict(date, time, dur) {
  const sm=timeToMin(time), em=sm+dur;
  const bookings=getBookingsForDate(date);
  for (const b of bookings) {
    const bs=timeToMin(b.time), be=bs+b.duration;
    if (sm<be && em>bs) return true;
  }
  return false;
}

function updateClient(phone, name) {
  const s = getSheet('Clients');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0]===phone) { s.getRange(i+1,2).setValue(name); return; }
  }
  s.appendRow([phone, name, '', '', '']);
}

const SERVICES = {
  manicure_strengthen: { name:'Маникюр с укреплением', price:1800 },
  extension: { name:'Наращивание', price:2500 },
  complex_design: { name:'Сложный дизайн', price:3200 },
  hardware_manicure: { name:'Аппаратный маникюр', price:1000 }
};

function book(p) {
  const {service, date, time, clientName, clientPhone} = p;
  const duration = parseInt(p.duration);
  if (!service||!date||!time||!duration||!clientName||!clientPhone) return {ok:false,error:'Все поля обязательны'};
  if (!/^\d{11}$/.test(clientPhone)) return {ok:false,error:'Формат телефона: 79991234567'};
  if (isBlacklisted(clientPhone)) return {ok:false,error:'Ваш номер заблокирован для записи'};
  if (!getOpenDatesArr().includes(date)) return {ok:false,error:'Дата недоступна'};
  if (!getSlotsArr(date).includes(time)) return {ok:false,error:'Время недоступно'};
  if (isConflict(date, time, duration)) return {ok:false,error:'Это время уже занято'};
  const sm=timeToMin(time), em=sm+duration;
  if (sm<480||em>1320) return {ok:false,error:'Время вне рабочих часов (08:00-22:00)'};
  const id = genId();
  const svc = SERVICES[service] || { name: service };
  getSheet('Bookings').appendRow([id, service, svc.name, date, time, duration, clientName, clientPhone, 'pending', new Date().toISOString()]);
  updateClient(clientPhone, clientName);
  return {ok:true, id, message:'Запись создана'};
}

function check(p) {
  const {clientPhone} = p;
  if (!clientPhone||!/^\d{11}$/.test(clientPhone)) return {ok:false,error:'Некорректный номер'};
  const data = getSheet('Bookings').getDataRange().getValues();
  const now = new Date();
  const bookings = [];
  for (let i=1; i<data.length; i++) {
    if (data[i][7]===clientPhone && new Date(data[i][3]+'T'+data[i][4])>=now && data[i][8]!=='cancelled') {
      bookings.push({
        service: data[i][1], serviceName: data[i][2], date: data[i][3],
        time: data[i][4], duration: data[i][5], status: data[i][8]
      });
    }
  }
  return {ok:true, bookings};
}

function getOpenDates() {
  return {ok:true, dates: getOpenDatesArr()};
}

function getSlots(p) {
  const {date} = p;
  const duration = parseInt(p.duration);
  if (!date||!duration) return {ok:false,error:'Укажите date и duration'};
  const raw = getSlotsArr(date);
  const avail = [];
  for (const slot of raw) {
    const sm=timeToMin(slot), em=sm+duration;
    if (em>1320) continue;
    if (!isConflict(date, slot, duration)) avail.push(slot);
  }
  return {ok:true, slots:avail};
}

function getRawSlots(p) {
  if (!p.date) return {ok:false,error:'Укажите date'};
  return {ok:true, slots: getSlotsArr(p.date)};
}

function openDate(p) {
  if (!p.date) return {ok:false,error:'Укажите date'};
  const s = getSheet('OpenDates');
  const arr = getOpenDatesArr();
  if (!arr.includes(p.date)) s.appendRow([p.date]);
  return {ok:true};
}

function closeDate(p) {
  if (!p.date) return {ok:false,error:'Укажите date'};
  const s = getSheet('OpenDates');
  const data = s.getDataRange().getValues();
  for (let i=data.length-1; i>=1; i--) if (data[i][0]===p.date) s.deleteRow(i+1);
  return {ok:true};
}

function openAllMonth(p) {
  const month = parseInt(p.month), year = parseInt(p.year);
  if (!month||!year) return {ok:false,error:'Укажите month и year'};
  const existing = getOpenDatesArr();
  const days = new Date(year, month, 0).getDate();
  const toAdd = [];
  for (let d=1; d<=days; d++) {
    const ds = year+'-'+String(month).padStart(2,'0')+'-'+String(d).padStart(2,'0');
    if (!existing.includes(ds)) toAdd.push([ds]);
  }
  if (toAdd.length) getSheet('OpenDates').getRange(getSheet('OpenDates').getLastRow()+1,1,toAdd.length,1).setValues(toAdd);
  return {ok:true, added: toAdd.length};
}

function closeAllMonth(p) {
  const month = parseInt(p.month), year = parseInt(p.year);
  if (!month||!year) return {ok:false,error:'Укажите month и year'};
  const s = getSheet('OpenDates');
  const data = s.getDataRange().getValues();
  for (let i=data.length-1; i>=1; i--) {
    const parts = String(data[i][0]).split('-');
    if (parseInt(parts[0])===year && parseInt(parts[1])===month) s.deleteRow(i+1);
  }
  return {ok:true};
}

function openSlot(p) {
  if (!p.date||!p.time) return {ok:false,error:'Укажите date и time'};
  const arr = getSlotsArr(p.date);
  if (arr.includes(p.time)) return {ok:false,error:'Слот уже открыт'};
  getSheet('OpenSlots').appendRow([p.date, p.time]);
  return {ok:true};
}

function closeSlot(p) {
  if (!p.date||!p.time) return {ok:false,error:'Укажите date и time'};
  const s = getSheet('OpenSlots');
  const data = s.getDataRange().getValues();
  for (let i=data.length-1; i>=1; i--) if (data[i][0]===p.date && data[i][1]===p.time) s.deleteRow(i+1);
  return {ok:true};
}

function getBookings(p) {
  const data = getSheet('Bookings').getDataRange().getValues();
  const filter = p.date || null;
  const result = [];
  for (let i=1; i<data.length; i++) {
    if (filter && data[i][3]!==filter) continue;
    if (data[i][0]) result.push({
      id:data[i][0], service:data[i][1], serviceName:data[i][2],
      date:data[i][3], time:data[i][4], duration:data[i][5],
      clientName:data[i][6], clientPhone:data[i][7], status:data[i][8]
    });
  }
  result.sort((a,b)=>(a.date+a.time).localeCompare(b.date+b.time));
  return {ok:true, bookings:result};
}

function updateStatus(id, st) {
  const s = getSheet('Bookings');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0]===id) { s.getRange(i+1,9).setValue(st); return true; }
  }
  return false;
}

function confirmBooking(p) { return p.id?(updateStatus(p.id,'confirmed')?{ok:true}:{ok:false,error:'Не найдена'}):{ok:false,error:'Укажите id'}; }
function cancelBooking(p) { return p.id?(updateStatus(p.id,'cancelled')?{ok:true}:{ok:false,error:'Не найдена'}):{ok:false,error:'Укажите id'}; }

function deleteBooking(p) {
  if (!p.id) return {ok:false,error:'Укажите id'};
  const s = getSheet('Bookings');
  const data = s.getDataRange().getValues();
  for (let i=data.length-1; i>=1; i--) if (data[i][0]===p.id) { s.deleteRow(i+1); return {ok:true}; }
  return {ok:false,error:'Не найдена'};
}

function moveBooking(p) {
  const {id, newDate, newTime} = p;
  if (!id||!newDate||!newTime) return {ok:false,error:'Укажите id, newDate, newTime'};
  const s = getSheet('Bookings');
  const data = s.getDataRange().getValues();
  let row=-1, dur=0;
  for (let i=1; i<data.length; i++) {
    if (data[i][0]===id) { row=i+1; dur=parseInt(data[i][5])||0; break; }
  }
  if (row<0) return {ok:false,error:'Не найдена'};
  if (!getSlotsArr(newDate).includes(newTime)) return {ok:false,error:'Слот не открыт'};
  if (isConflict(newDate, newTime, dur)) return {ok:false,error:'Время занято'};
  s.getRange(row,4).setValue(newDate);
  s.getRange(row,5).setValue(newTime);
  return {ok:true};
}

function blacklistAdd(p) {
  const phone = p.phone, reason = p.reason || '';
  if (!phone||!/^\d{11}$/.test(phone)) return {ok:false,error:'Некорректный номер'};
  const s = getSheet('Blacklist');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0]===phone) return {ok:false,error:'Уже в чёрном списке'};
  s.appendRow([phone, reason, new Date().toISOString()]);
  return {ok:true};
}

function blacklistRemove(p) {
  const s = getSheet('Blacklist');
  const data = s.getDataRange().getValues();
  for (let i=data.length-1; i>=1; i--) if (data[i][0]===p.phone) { s.deleteRow(i+1); return {ok:true}; }
  return {ok:false,error:'Не найден'};
}

function getBlacklist() {
  const data = getSheet('Blacklist').getDataRange().getValues();
  const list = [];
  for (let i=1; i<data.length; i++) if (data[i][0]) list.push({phone:data[i][0], reason:data[i][1]});
  return {ok:true, list};
}

function clientHistory(p) {
  if (!p.clientPhone||!/^\d{11}$/.test(p.clientPhone)) return {ok:false,error:'Некорректный номер'};
  const data = getSheet('Bookings').getDataRange().getValues();
  const hist = [];
  for (let i=1; i<data.length; i++) {
    if (data[i][7]===p.clientPhone) hist.push({
      service:data[i][1], serviceName:data[i][2], date:data[i][3], time:data[i][4], status:data[i][8]
    });
  }
  hist.sort((a,b)=>(b.date+b.time).localeCompare(a.date+a.time));
  return {ok:true, history:hist};
}

function saveCost(p) {
  const {bookingId, cost, note} = p;
  if (!bookingId||isNaN(parseFloat(cost))) return {ok:false,error:'Укажите bookingId и cost'};
  getSheet('Costs').appendRow([bookingId, parseFloat(cost), note||'', new Date().toISOString()]);
  return {ok:true};
}

function stats(p) {
  const {startDate, endDate} = p;
  if (!startDate||!endDate) return {ok:false,error:'Укажите startDate и endDate'};
  const bs = getSheet('Bookings').getDataRange().getValues();
  const cs = getSheet('Costs').getDataRange().getValues();
  const costMap = {};
  for (let i=1; i<cs.length; i++) costMap[cs[i][0]] = parseFloat(cs[i][1])||0;
  let rev=0, cost=0;
  const clients=new Set(), details=[];
  for (let i=1; i<bs.length; i++) {
    if (bs[i][3]>=startDate && bs[i][3]<=endDate && bs[i][8]!=='cancelled') {
      const svc = SERVICES[bs[i][1]];
      const pr = svc ? svc.price : 0;
      const c = costMap[bs[i][0]]||0;
      rev+=pr; cost+=c;
      clients.add(bs[i][7]);
      details.push({date:bs[i][3],time:bs[i][4],clientName:bs[i][6],serviceName:bs[i][2],price:pr,cost:c,profit:pr-c});
    }
  }
  const cnt=details.length;
  return {ok:true, revenue:rev, totalCost:cost, profit:rev-cost, clientCount:clients.size, avgCheck:cnt?Math.round(rev/cnt):0, details};
}

function crmSearch(p) {
  const q = (p.query||'').toLowerCase().trim();
  if (!q) return {ok:false,error:'Введите запрос'};
  const bs = getSheet('Bookings').getDataRange().getValues();
  const cs = getSheet('Clients').getDataRange().getValues();
  const map = {};
  for (let i=1; i<bs.length; i++) {
    const ph=bs[i][7], nm=(bs[i][6]||'').toLowerCase();
    if (!ph) continue;
    if (ph.includes(q)||nm.includes(q)) {
      if (!map[ph]) map[ph]={clientPhone:ph,clientName:bs[i][6],bookingCount:0,lastVisit:'',tags:[],note:''};
      map[ph].bookingCount++;
      if (!map[ph].lastVisit||bs[i][3]>map[ph].lastVisit) map[ph].lastVisit=bs[i][3];
    }
  }
  for (let i=1; i<cs.length; i++) {
    const ph=cs[i][0];
    if (map[ph]) { map[ph].note=cs[i][2]||''; map[ph].tags=(cs[i][3]||'').split(',').filter(Boolean); }
  }
  return {ok:true,clients:Object.values(map)};
}

function saveClientNote(p) {
  const {clientPhone, note} = p;
  if (!clientPhone||!/^\d{11}$/.test(clientPhone)||!note) return {ok:false,error:'Телефон и заметка'};
  const s = getSheet('Clients');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0]===clientPhone) { s.getRange(i+1,3).setValue(note); return {ok:true}; }
  }
  s.appendRow([clientPhone,'',note,'','']);
  return {ok:true};
}