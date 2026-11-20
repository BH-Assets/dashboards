"use strict";
/* ---------------- Rate sheet (from BHESC_Rental_Rates.docx) ---------------- */
const RATE_SHEET = [["121","John Deere 700K Crawler",635.0,1920.0,5200.0],["124","Kubota SVL 75-2 Skid Steer",385.0,1035.0,2205.0],["128","Kubota SVL 75-2 Skid Steer",385.0,1035.0,2205.0],["129","Cat 255 Skid Steer",395.0,1105.0,2345.0],["130","Cat 305 Mini Excavator",425.0,1100.0,2595.0],["131","Cat 306 Mini Excavator",445.0,1185.0,2695.0],["132","Cat 255 Skid Steer",395.0,1105.0,2345.0],["133","Cat 255 Skid Steer (Enclosed Cab)",425.0,1175.0,2445.0],["209","GEHL DL-7L Lift",500.0,1150.0,2450.0],["305","Dump Truck Single Axle 6YD",385.0,1100.0,3400.0],["306","Dump Truck Single Axle 6YD",385.0,1100.0,3400.0],["405","18\u2019 Equipment Trailer",50.0,250.0,800.0],["406","18\u2019 Deckover Trailer",50.0,250.0,800.0],["410","Conex \u2013",null,null,100.0],["411","Conex \u2013",null,null,100.0],["412","Conex - Open",null,null,100.0],["413","Conex \u2013 Eric Studdard",null,null,100.0],["414","Conex \u2013 Eric Studdard",null,null,100.0],["415","Conex \u2013 Beau Newman",null,null,100.0],["416","Conex \u2013",null,null,100.0],["417","Conex \u2013 Self Preform",null,null,100.0],["418","Conex \u2013 Heath Whorton",null,null,100.0],["419","Conex \u2013 BH Residential",null,null,100.0],["420","Pressure Washer Trailer",200.0,1000.0,1600.0],["503","Terramite Broom Sweeper",315.0,835.0,1685.0],["504","Terramite Broom Sweeper",315.0,835.0,1685.0],["508","T12 Ride on Scrubber",315.0,840.0,1875.0],["510","Kubota 1140 UTV",150.0,325.0,1050.0],["511","Kubota 1140 UTV",150.0,325.0,1050.0],["512","Multiquip Wheeled Georgia Buggy",130.0,370.0,795.0],["513","Broce BW260 Broom Sweeper",315.0,835.0,1685.0],["905","Equipment Maintenance Station",null,null,500.0],["906","Portable Water Tank",125.0,240.0,580.0],["908","BB121 Grader Box",240.0,630.0,2400.0],["909","Generac Generator 6500w",75.0,240.0,400.0],["910","Generac Generator 6500w",75.0,240.0,400.0],["912","Generac Generator 6500w \u2013 Electric Start",75.0,240.0,400.0],["913","Generac Generator 6500w",75.0,240.0,400.0],["917","Pressure Washer 3500psi",75.0,240.0,400.0],["920","Plate Tamp",95.0,270.0,495.0],["921","Plate Tamp",95.0,270.0,495.0],["922","Plate Tamp",95.0,270.0,495.0],["923","Plate Tamp",95.0,270.0,495.0],["924","BH Signs- Crane Lighted Sign",null,null,620.0],["925","Generac Generator 6500w",75.0,240.0,400.0],["926","Generac Generator 6500w",75.0,240.0,400.0],["927","Jumping Jack",105.0,275.0,495.0],["928","Pressure Washer 3500psi",75.0,240.0,400.0],["930","Pressure Washer 3500psi",75.0,240.0,400.0],["931","Skid Steer Auger Attachment",125.0,340.0,810.0],["931-10","Skid Steer Auger Bit 10\u201d",42.0,125.0,350.0],["931-12","Skid Steer Auger Bit 12\u201d",45.0,150.0,370.0],["931-20","Skid Steer Auger Bit 20\u201d",50.0,175.0,385.0],["931-24","Skid Steer Auger Bit 24\u201d",55.0,185.0,420.0],["932","Plate Tamp",95.0,270.0,495.0],["933","Jumping Jack",105.0,275.0,495.0],["934","Light Tower",125.0,275.0,565.0],["935","Light Tower",125.0,275.0,565.0],["936","Light Tower",125.0,275.0,565.0],["937","Cat Trencher T9B",175.0,420.0,1015.0],["939","Firman Generator",75.0,240.0,400.0],["940","Skip Pan",90.0,240.0,600.0],["941","Skip Pan",90.0,240.0,600.0],["943","36\u201d Concrete Finisher",75.0,240.0,400.0],["946","Plate Tamp",95.0,270.0,495.0],["950","Air Chipping Hammer - 60 lbs",50.0,180.0,375.0],["951","Generac Generator 6500w \u2013 Electric Start",75.0,240.0,400.0],["952","Bosch 65lbs Demo Hammer",80.0,245.0,500.0],["955","Job Box 1",null,null,250.0],["956","Job Box 2",null,null,250.0],["957","Job Box 3",null,null,250.0],["958","Job Box 4",null,null,250.0],["959","48\u201d Pallet Fork Skid steer Attachment",50.0,140.0,265.0],["960","Husqvarna Demo Saw",60.0,168.0,450.0],["961","Husqvarna Demo Saw",60.0,168.0,450.0],["962","48\u201d Pallet Fork Skid steer Attachment",50.0,140.0,265.0],["963","\u00bd Yard Concrete Bucket",60.0,90.0,360.0],["964","\u00be Yard Concrete Bucket",85.0,135.0,405.0],["965","1 Yard Concrete Bucket",95.0,205.0,450.0],["966","2 Yard Concrete Bucket",115.0,265.0,585.0],["967","3 Yard Concrete Bucket",150.0,350.0,765.0],["969","Jumping Jack",105.0,275.0,495.0],["970","48\u201d Pallet Fork Skid steer Attachment",50.0,140.0,265.0],["971","Bosch 35lb Demo Hammer",65.0,200.0,450.0],["972","Bosch 25lb Demo Hammer",60.0,175.0,425.0],["973","Grapple Rake",112.0,275.0,580.0],["974","Generac Light Tower",125.0,275.0,565.0],["976","Drum Fan 36\u201d",50.0,125.0,350.0],["977","Drum Fan 36\u201d",50.0,125.0,350.0],["978","Drum Fan 36\u201d",50.0,125.0,350.0],["979","Diesel Torpedo Heater",50.0,125.0,350.0],["980","Diesel Torpedo Heater",50.0,125.0,350.0],["981","Diesel Torpedo Heater",50.0,125.0,350.0],["982","Kerosene Torpedo Heater",50.0,125.0,350.0],["983","Kerosene Torpedo Heater",50.0,125.0,350.0],["984","Kerosene Torpedo Heater",50.0,125.0,350.0],["985","Pressure Washer 4000 PSI",75.0,240.0,400.0],["986","Stihl TS-420 Demo Saw",85.0,225.0,450.0]];

const KEY = "bh-rentals-v1";
const PICKUPS = ["Customer Pickup","Customer Returned","BHESC Delivered","BHESC Pickup"];
const DAY = 86400000;
const ICONS = {
  dashboard:'<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
  orders:'<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h4"/>',
  request:'<path d="M9 4h6a1 1 0 0 1 1 1v1H8V5a1 1 0 0 1 1-1Z"/><path d="M8 5H6a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1h-2"/><path d="M12 10v6M9 13h6"/>',
  "my-requests":'<path d="M9 4h6a1 1 0 0 1 1 1v1H8V5a1 1 0 0 1 1-1Z"/><path d="M8 5H6a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1h-2"/><path d="m9 13 2 2 4-4"/>',
  requests:'<path d="M3 13h5l2 3h4l2-3h5"/><path d="M5 5h14l2 8v6a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-6Z"/>',
  items:'<path d="M21 8 12 3 3 8l9 5 9-5Z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/>',
  maintenance:'<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.4-.6-.6-2.4 2.5-2.5Z"/>',
  locations:'<path d="M12 21s7-6.2 7-12a7 7 0 1 0-14 0c0 5.8 7 12 7 12Z"/><circle cx="12" cy="9" r="2.5"/>',
  customers:'<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
  reports:'<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  setup:'<path d="M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M20 18h0"/><circle cx="16" cy="6" r="2"/><circle cx="10" cy="12" r="2"/><circle cx="18" cy="18" r="2"/>',
  checkout:'<path d="M9 11V6a3 3 0 0 1 6 0v5"/><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M12 15v2"/>',
  optimizer:'<path d="M12 3v9l6.4 6.4A9 9 0 1 1 12 3Z"/><path d="M15 3.5a9 9 0 0 1 5.5 5.5H15Z"/>',
  stock:'<rect x="3" y="12" width="8" height="8" rx="1"/><rect x="13" y="12" width="8" height="8" rx="1"/><rect x="8" y="4" width="8" height="8" rx="1"/>',
  inventory:'<path d="M3 7h18v13H3zM3 7l2-4h14l2 4M9 11h6"/>',
  workshop:'<path d="M3 21h18M5 21V10l7-5 7 5v11"/><path d="M9 21v-6h6v6"/>',
  bundles:'<path d="m12 2 9 5-9 5-9-5 9-5Z"/><path d="m3 12 9 5 9-5M3 17l9 5 9-5"/>',
  groups:'<circle cx="12" cy="5" r="2.5"/><circle cx="5" cy="19" r="2.5"/><circle cx="19" cy="19" r="2.5"/><path d="M12 7.5v4M12 11.5 5.8 16.8M12 11.5l6.2 5.3"/>',
  arrow:'<path d="M5 12h14M13 6l6 6-6 6"/>',
  plus:'<path d="M12 5v14M5 12h14"/>',
  menu:'<path d="M4 7h16M4 12h16M4 17h16"/>',
  arrowback:'<circle cx="12" cy="12" r="9"/><path d="M14 8l-4 4 4 4M10 12h6"/>',
  moon:'<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
  logout:'<path d="M15 4h4v16h-4M10 8l-4 4 4 4M6 12h10"/>',
  download:'<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>',
  printer:'<path d="M7 9V3h10v6M7 17H4v-7h16v7h-3"/><rect x="7" y="14" width="10" height="7"/>',
  clip:'<path d="m20 11-8.5 8.5a5 5 0 0 1-7-7L13 4a3.5 3.5 0 0 1 5 5l-8.4 8.4a2 2 0 0 1-2.8-2.8L14.5 7"/>',
  camera:'<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
  edit:'<path d="M4 20h4L19 9l-4-4L4 16v4Z"/><path d="m13.5 6.5 4 4"/>',
  copy:'<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h3"/>',
  trash:'<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
  available:'<rect x="3" y="4" width="18" height="16" rx="2"/><path d="m8 12 3 3 5-6"/>',
  search:'<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
};
const ico = (n, sw=2) => (ICONS[n] ||= ICONS[{"my-reports":"reports"}[n]] || "", icoRaw(n, sw));
const icoRaw = (n, sw=2) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round">${ICONS[n]}</svg>`;
const ITEM_SUB = [["items","Assets","items"],["asset-stock","Asset Stock","stock"],["inventory","Inventory","inventory"],["groups","Groups","groups"]];
const ITEM_KEYS = ITEM_SUB.map(x=>x[0]);
const NAV = [["dashboard","Dashboard"],["orders","Orders"],["checkout","Check In/Out"],["items","Items"],["maintenance","Maintenance"],["locations","Locations"],["customers","Customers"],["reports","Reports"],["setup","Setup"]];

/* ---------------- Helpers ---------------- */
const $ = s => document.querySelector(s);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const money = n => (n == null || isNaN(n)) ? "—" : "$" + Number(n).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const money0 = n => "$" + Math.round(n||0).toLocaleString("en-US");
const uid = p => p + Math.random().toString(36).slice(2,9);
const pad = n => String(n).padStart(2,"0");
const toLocalInput = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
const parseD = s => s ? new Date(/^\d{4}-\d{2}-\d{2}$/.test(s) ? s+"T00:00" : s) : null;
const fmtDate = s => { const d = parseD(s); return d ? `${pad(d.getMonth()+1)}/${pad(d.getDate())}/${d.getFullYear()}` : "—"; };
const fmtDT = s => { const d = parseD(s); return d ? `${fmtDate(s)}<br><span class="muted">${pad(d.getHours())}:${pad(d.getMinutes())}</span>` : "—"; };
const now = () => new Date();

const SERIES = {"1":"100 - Excavation","2":"200 - Lifts","3":"300 - Trucks","4":"400 - Trailers & Storage","5":"500 - Site Equipment","6":"600 Series","7":"700 Series","8":"800 Series","9":"900 - Small Equipment & Tools"};
function seriesGroup(id){ const m = /^([1-9])\d\d(?!\d)/.exec(String(id||"").trim()); return m ? SERIES[m[1]] : null; }
function seriesDesc(name){ const m = /^(\d)00/.exec(name); return m ? `Asset numbers ${m[1]}00–${m[1]}99` : ""; }
function categorize(name, id){ const s = seriesGroup(id); if (s) return s; return categorizeByName(name); }
function categorizeByName(name){
  const n = name.toLowerCase();
  if (/crawler|excavator|skid steer(?! auger)(?! attach)|trencher/.test(n) && !/attachment|auger|fork/.test(n)) return "Earthmoving";
  if (/dump truck|trailer(?! *$)/.test(n) && !/pressure washer trailer/.test(n)) return "Trucks & Trailers";
  if (/gehl|lift/.test(n)) return "Lifts";
  if (/conex|job box/.test(n)) return "Storage";
  if (/auger|pallet fork|grapple|grader box|fork mounted/.test(n)) return "Attachments";
  if (/generator|light tower|heater|fan/.test(n)) return "Power, Light & Climate";
  if (/tamp|jumping jack|concrete|skip pan|finisher|buggy/.test(n)) return "Compaction & Concrete";
  if (/saw|hammer|pressure washer 3500|pressure washer 4000/.test(n)) return "Tools";
  return "Site Services";
}

function seed(){
  const items = RATE_SHEET.map(([id,name,day,week,month]) => ({id,name:name.replace(/\s*[–-]\s*$/,""),category:categorize(name,id),day,week,month,status:"Active",serial:"",notes:""}));
  items.push({id:"975",name:"Fork Mounted Attachment",category:seriesGroup("975"),day:null,week:null,month:450,status:"Active",serial:"",notes:"Appears on order 2088 but is not on the rate sheet. Confirm the description and rates."});
  items.sort((a,b)=>a.id.localeCompare(b.id,undefined,{numeric:true}));
  const people = [["c-studdard","Eric Studdard"],["c-mickle","Wendy Mickle"],["c-roberts","Coleby Roberts"],["c-whorton","Heath Whorton"],["c-patterson","Mike Patterson"],["c-sanders","Garrett Sanders"],["c-lomenick","Blake Lomenick"]];
  const customers = people.map(([id,name]) => ({id,name,company:"Bailey Harris",phone:"",email:"",notes:""}));
  const jobs = [["l-2504","2504-00","Additions"],["l-au001","25AU001","Arbor Wood"],["l-au002","25AU002","GL Pack"],["l-au013","25AU013","RFID Facility"],["l-au018","25AU018","EAH MRI"],["l-au010","25AU010","AU Comer"],["l-yard","","BHESC Shop Yard"]];
  const locations = jobs.map(([id,code,name]) => ({id,code,name,address:"",contact:"",notes:"",...(id==="l-yard"?{isYard:true}:{})}));
  const mk = (no,cust,loc,itemId,out,ret,opts={}) => {
    const it = items.find(i=>i.id===itemId);
    return {no,customerId:cust,locationId:loc,lines:[{itemId,day:it.day,week:it.week,month:it.month}],out,ret,status:"Rented Out",
      description:opts.description||"",pickup:"--",delivery:opts.delivery||{name:"",rate:0,trips:0},charges:opts.charges||[],adjustment:opts.adjustment||0,
      returnedAt:null,notes:"",created:out,log:[{at:out,text:"Order created and rented out"}]};
  };
  const orders = [
    mk(2082,"c-studdard","l-2504","964","2026-09-30T06:00","2026-10-28T06:00",{delivery:{name:"Local",rate:150,trips:2}}),
    mk(2084,"c-mickle","l-au001","419","2026-09-30T08:00","2026-10-28T07:59",{description:"419 - Conex"}),
    mk(2087,"c-roberts","l-au002","510","2026-10-01T07:00","2026-10-29T07:00"),
    mk(2085,"c-whorton","l-au013","418","2026-10-01T09:00","2026-10-29T09:00"),
    mk(2090,"c-patterson","l-au018","410","2026-10-05T08:30","2026-11-02T08:30"),
    mk(2088,"c-studdard","l-2504","975","2026-10-05T09:04","2026-11-02T09:04"),
    mk(2091,"c-sanders","l-au010","986","2026-10-06T09:00","2026-11-03T09:00"),
    mk(2095,"c-lomenick","l-au001","133","2026-10-12T13:00","2026-11-09T13:00",{delivery:{name:"<125 Miles",rate:260,trips:2},adjustment:50}),
  ];
  return {
    v:1, nextOrder:2096,
    settings:{company:"BH Equipment",contact:"John Barron",phone:"334.750.1753",email:"jbarron@baileyharris.com",cycle:28,
      delivery:[{name:"Local",rate:150},{name:"<125 Miles",rate:260},{name:"<200 Miles",rate:380},{name:">200 Miles",rate:500}],
      charges:[{name:"Mechanic service",rate:115,unit:"hr"},{name:"Diesel delivered",rate:8.5,unit:"gal"},{name:"Diesel on return",rate:8.5,unit:"gal"},{name:"Gas on return",rate:7.5,unit:"gal"},{name:"Excess dirt",rate:50,unit:"each"},{name:"Concrete splatter",rate:100,unit:"each"},{name:"F150 pickup truck",rate:900,unit:"month"}],
      policy:"All projects are responsible for routine service and cleaning of the undercarriage and equipment as required, and for returning rented equipment promptly and in good order."},
    items, customers, locations, orders, maintenance:[],
  };
}

let S;
S = null; /* cloud.js loads the real data from the Equipment service */
if (!S || !S.items) S = seed();
migrate(S);
function migrate(d){
  d.users ||= [{id:"u-admin",name:d.settings.contact||"Administrator",email:d.settings.email||"",role:"Admin",active:true,hash:null,title:"",phone:d.settings.phone||""}];
  d.settings.terms ||= "Net 30"; if (d.settings.taxRate==null) d.settings.taxRate = 0;
  // find the shop yard even if it was renamed; never fall back to a job site
  if (!d.locations.some(l=>l.isYard)){
    const old = d.locations.find(l=>l.id==="l-yard") || d.locations.find(l=>(l.code||"").toUpperCase()==="YARD") || d.locations.find(l=>/yard/i.test((l.code||"")+" "+l.name));
    if (old) Object.assign(old, {isYard:true, code:"", name:"BHESC Shop Yard"});
    else d.locations.push({id:"l-yard-"+Math.random().toString(36).slice(2,7), code:"", name:"BHESC Shop Yard", address:"", contact:"", notes:"", isYard:true});
  }
  // repair (once): merge duplicate yard records and send assets wrongly homed at a job site back to the yard
  if (!d.yardRepair1){
    const yards = d.locations.filter(l=>l.isYard || l.id===d.locations.find(x=>x.isYard).id);
    if (yards.length > 1){
      const keep = yards.find(l=>l.isYard) || yards[0];
      yards.forEach(l=>{ if (l===keep) return; ["address","contact","notes"].forEach(k=>{ if (!keep[k] && l[k]) keep[k] = l[k]; }); });
      Object.assign(keep, {isYard:true, code:"", name:"BHESC Shop Yard"});
      d.locations = d.locations.filter(l=>!yards.includes(l) || l===keep);
    }
    const yid = d.locations.find(l=>l.isYard).id, label = l => l ? (l.code ? `${l.code} - ${l.name}` : l.name) : "";
    d.items.forEach(i=>{
      const l = d.locations.find(x=>x.id===i.locationId);
      if (!l || l.isYard) return;
      const chosen = (i.log||[]).some(e=>/Location .*→/.test(e.text||"") && (e.text||"").includes("→ "+label(l)));
      if (!chosen) i.locationId = yid;
    });
    d.yardRepair1 = true;
  }
  const yard = d.locations.find(l=>l.isYard).id; d.items.forEach(i=>{ if (!i.locationId || !d.locations.some(l=>l.id===i.locationId)) i.locationId = yard; });
  d.orders.forEach(o=>{ if (o.status==="Reserved") o.status = "Booked";
    if (!PICKUPS.includes(o.pickup)) o.pickup = ({"Requested":"BHESC Pickup","Scheduled":"BHESC Pickup","Picked Up":"BHESC Pickup"})[o.pickup] || ""; });
  d.custody ||= []; d.nextCheckout ||= 1001; d.stock ||= [];
  if (!d.inventory) d.inventory = [
    {id:uid("i"),name:"Diesel fuel",sku:"FUEL-DSL",unit:"gal",qty:0,reorder:0,unitCost:null,price:8.5,location:"Yard",vendor:"",log:[]},
    {id:uid("i"),name:"Gasoline",sku:"FUEL-GAS",unit:"gal",qty:0,reorder:0,unitCost:null,price:7.5,location:"Yard",vendor:"",log:[]}];
  d.bundles ||= [];
  d.groups ||= [];
  if (!d.groupsByNumber){
    d.items.forEach(i=>{ const s = seriesGroup(i.id); if (s) i.category = s; });
    const used = new Set(d.items.map(i=>i.category).concat((d.stock||[]).map(s=>s.category)));
    const series = Object.values(SERIES).filter(n=>d.items.some(i=>i.category===n));
    d.groups = series.map(n=>({name:n, desc:(d.groups.find(g=>g.name===n)||{}).desc||seriesDesc(n)}))
      .concat(d.groups.filter(g=>!series.includes(g.name) && used.has(g.name)));
    d.groupsByNumber = true;
  }
  [...new Set(d.items.map(i=>i.category))].forEach(c=>{ if (!d.groups.some(g=>g.name===c)) d.groups.push({name:c,desc:""}); });
  d.maintenance.forEach(m=>{ m.stage ||= m.status==="Open"?"Queued":"Done"; m.parts ||= []; });
  // service requests (GCSS-MC style): SR numbers, lifecycle, tasks with labor and parts
  d.settings.nextSR ||= 1001;
  d.requests ||= []; d.settings.nextRQ ||= 1001;
  // one time: every inventory item is kept at the Equipment Shop
  if (!d.invToShop1){ const shop = invHomeId(d.locations||[]); if (shop) (d.inventory||[]).forEach(i=>{ i.locationId = shop; }); d.invToShop1 = true; }
  // Rental Coordinator and Mechanic roles were removed: those users become Field Crew
  (d.users||[]).forEach(u=>{ if (u.role==="Rental Coordinator" || u.role==="Mechanic"){ u.formerRole = u.role; u.role = "Field Crew"; } });
  d.maintenance.slice().sort((a,b)=>(a.date||"").localeCompare(b.date||"")).forEach(m=>{ if (!m.no) m.no = d.settings.nextSR++; });
  const typeMap = {"Repair":"Corrective maintenance","Breakdown call":"Corrective maintenance","Preventive":"Preventive maintenance","Cleaning":"Preventive maintenance","Inspection":"Inspection"};
  const stageMap = {"Queued":"Accepted","In progress":"In work","Waiting on parts":"Awaiting parts","Done":"Closed"};
  d.maintenance.forEach(m=>{ if (m.srV) return; m.srV = 1;
    m.type = typeMap[m.type] || m.type || "Corrective maintenance";
    m.stage = m.status==="Completed" ? "Closed" : (stageMap[m.stage] || m.stage || "Accepted");
    m.acceptedAt ||= m.date; m.codes ||= []; m.reportedBy ||= ""; m.log ||= [];
    m.opStatus ||= m.status==="Open" ? (m.severity==="High" ? "Deadlined" : m.severity==="Medium" ? "Degraded" : "Operational") : "Operational";
    if (m.opStatus==="Deadlined" && m.status==="Open") m.deadlineAt ||= m.date+"T00:00";
    if (!m.tasks){ const inv = id => (d.inventory||[]).find(i=>i.id===id);
      const t = {id:uid("t"), subject:(m.desc||"Work performed").split("\n")[0].slice(0,90), code:"", status:m.status==="Open"?"In work":"Complete", owner:m.tech||"", planned:"", labor:[], parts:[]};
      if (+m.hours) t.labor.push({id:uid("l"), who:m.tech||"", date:m.doneAt||m.date, hours:+m.hours, note:""});
      (m.parts||[]).forEach(p=>t.parts.push({id:uid("p"), invId:p.invId, name:inv(p.invId)?.name||"Part", partNo:inv(p.invId)?.sku||"", vendor:"", po:"", qty:+p.qty||0, cost:+p.cost||0, status:"Installed", taken:+p.qty||0}));
      m.parts = []; m.tasks = [t]; }
    if (m.status==="Completed" && !m.qc) m.qc = {by:m.tech||"", at:(m.doneAt||m.date)+"T00:00", note:"Closed before QC tracking", passed:true};
    m.log.push({at:(m.date||"")+"T00:00", by:"", text:"Converted from the old work order format"}); });
  // stock and inventory: "Kept at" is a location from the system (old typed text is matched by name or code, otherwise the yard)
  const locKey = l => [(l.code ? `${l.code} - ${l.name}` : l.name), l.name, l.code].filter(Boolean).map(x=>x.toLowerCase());
  [...(d.stock||[]), ...(d.inventory||[])].forEach(x=>{ if (x.locationId && d.locations.some(l=>l.id===x.locationId)) return;
    const t = String(x.location||"").trim().toLowerCase(), m = t && d.locations.find(l=>locKey(l).includes(t));
    x.locationId = m ? m.id : d.locations.find(l=>l.isYard).id; });
  return d;
}
let fileHandle = null, storageFull = false;
/* Browser storage: localStorage holds only ~5 MB, so the full data (photos and files) is also kept in IndexedDB, which holds far more. */
const IDB = {
  db: null,
  open(){ return new Promise((res,rej)=>{ if (!window.indexedDB) return rej(new Error("no indexedDB")); const r = indexedDB.open("bh-rentals", 1); r.onupgradeneeded = () => r.result.createObjectStore("kv"); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  async get(k){ this.db ||= await this.open(); return new Promise((res,rej)=>{ const q = this.db.transaction("kv").objectStore("kv").get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
  async set(k, v){ this.db ||= await this.open(); return new Promise((res,rej)=>{ const t = this.db.transaction("kv","readwrite"); t.objectStore("kv").put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); }); }
};
let idbTimer = null, idbWarned = false;
function idbSave(json){ IDB.set(KEY, json || JSON.stringify(S)).then(()=>{ idbWarned = false; }).catch(()=>{
  if (storageFull && !fileHandle && !idbWarned && typeof toast==="function"){ idbWarned = true; toast("This browser couldn't keep your latest changes. Use Setup → Save data file so nothing is lost."); } }); }
function save(){
  [...new Set(S.items.map(i=>i.category))].forEach(c=>{ if (!S.groups.some(g=>g.name===c)) S.groups.push({name:c,desc:seriesDesc(c)}); });
  S._savedAt = Date.now();
  const json = JSON.stringify(S);
  try { localStorage.setItem(KEY, json); storageFull = false; } catch(e) { storageFull = true; }
  idbSave(json);
  if (fileHandle) writeFile().catch(()=>{});
  updateBell();
}

/* ---------------- Pricing ---------------- */
const wall = s => { const d = parseD(s); return Date.UTC(d.getFullYear(),d.getMonth(),d.getDate(),d.getHours(),d.getMinutes()); };
function rentalDays(out, ret){ return Math.max(1, Math.ceil(((wall(ret) - wall(out)) / DAY) - 0.01)); }
function priceLine(line, out, ret, cycLen){
  const days = rentalDays(out, ret), cyc = cycLen || +S.settings.cycle || 28;
  const opts = [[1,line.day,"day"],[7,line.week,"wk"],[cyc,line.month,"mo"]].filter(o => o[1] != null && o[1] !== "" && !isNaN(o[1]));
  if (!opts.length) return {total:0,days,label:"No rate"};
  const f = new Array(days+1).fill(Infinity), pick = new Array(days+1).fill(null); f[0]=0;
  for (let n=1;n<=days;n++) for (const o of opts){ const c = +o[1] + f[Math.max(0,n-o[0])]; if (c < f[n]) { f[n]=c; pick[n]=o; } }
  const cnt = {}; let n = days;
  while (n>0){ const o = pick[n]; cnt[o[2]] = (cnt[o[2]]||0)+1; n = Math.max(0,n-o[0]); }
  const label = ["mo","wk","day"].filter(k=>cnt[k]).map(k=>`${cnt[k]} ${k}${cnt[k]>1&&k!=="day"?"s":k==="day"&&cnt[k]>1?"s":""}`).join(" + ");
  return {total:f[days],days,label};
}
/* ---------------- Cycle billing: bill every cycle while the equipment stays on the job ---------------- */
const cycleDays = () => +S.settings.cycle || 28;
function cycleName(o){ return o.cycle==="month" ? "monthly, on the same date each month" : o.cycle ? `every ${cycleDays()} days` : ""; }
function cycleAt(o, n){ const d = parseD(o.out), x = new Date(d);
  if (o.cycle==="month"){ const day = d.getDate(); x.setDate(1); x.setMonth(x.getMonth()+n); x.setDate(Math.min(day, new Date(x.getFullYear(), x.getMonth()+1, 0).getDate())); return x; }
  x.setDate(x.getDate() + n*cycleDays()); return x; }
/* the billing periods from rent-out to `end`; the last one may be partial */
function cycleSegs(o, end){
  const e = parseD(end), out = [];
  if (!o.cycle || !(e > parseD(o.out))) return out;
  for (let n = 0; n < 600; n++){
    const a = cycleAt(o, n), b = cycleAt(o, n+1), full = b <= e;
    out.push({n:n+1, from:toLocalInput(a), to:toLocalInput(full ? b : e), full, len:rentalDays(toLocalInput(a), toLocalInput(b))});
    if (!full || +b === +e) break;
  }
  return out;
}
/* how far the cycles run: to the return once it's back, otherwise to the booked return or today, whichever is later */
function cycEnd(o){ if (o.status==="Returned") return orderRentalEnd(o); const n = toLocalInput(now()); return parseD(o.ret) > now() ? o.ret : n; }
const billFor = (o, n) => (o.bills||[]).find(b=>b.n===n);
function piecesRental(l, pieces){
  let total = 0, days = 0; const labels = [];
  pieces.forEach(pc=>{ const p = priceLine(l, pc.from, pc.to, pc.len); total += p.total; days += p.days; labels.push(p.label); });
  const same = labels.length > 1 && labels.every(x=>x===labels[0]);
  const fulls = pieces.filter(p=>p.full !== false && rentalDays(p.from,p.to)===p.len), part = pieces.filter(p=>!fulls.includes(p));
  const label = pieces.length===1 ? labels[0] : same ? `${labels.length} × ${labels[0]}` : [fulls.length?`${fulls.length} × ${priceLine(l,fulls[0].from,fulls[0].to,fulls[0].len).label}`:"", ...part.map(p=>priceLine(l,p.from,p.to,p.len).label)].filter(Boolean).join(" + ");
  return {total, days, label};
}
function lineRental(o, l, end){
  const segs = cycleSegs(o, end);
  return segs.length ? piecesRental(l, segs) : priceLine(l, o.out, end);
}
/* what's left for the final invoice: every cycle, minus what earlier cycle invoices already covered */
function finalPieces(o, end){
  return cycleSegs(o, end).flatMap(s=>{ const b = billFor(o, s.n); if (!b) return [s]; if (parseD(b.to) >= parseD(s.to)) return []; return [{from:b.to, to:s.to, len:s.len, full:false}]; });
}
function billKeys(o){
  const k = [];
  if ((+o.delivery?.rate||0) * (+o.delivery?.trips||0)) k.push("delivery");
  (o.sales||[]).forEach((_,i)=>k.push("sale:"+i));
  (o.charges||[]).forEach((_,i)=>k.push("charge:"+i));
  if (+o.adjustment) k.push("adj");
  return k;
}
function unbilledKeys(o){ const done = new Set((o.bills||[]).flatMap(b=>b.keys||[])); return billKeys(o).filter(k=>!done.has(k)); }
function cycleDue(o){ return o.cycle && o.status==="Rented Out" ? cycleSegs(o, cycEnd(o)).filter(s=>parseD(s.to) <= now() && !billFor(o, s.n)) : []; }
function cyclePanel(o){
  if (!o.cycle) return "";
  if (!["Rented Out","Returned"].includes(o.status)) return `<div class="infobox">Cycle billing, ${esc(cycleName(o))}. The first cycle starts when it's rented out.</div>`;
  const segs = cycleSegs(o, cycEnd(o)), back = o.status==="Returned", nowD = now();
  const billed = (o.bills||[]).reduce((s,b)=>s+(+b.total||0),0);
  const rows = segs.map(s=>{ const b = billFor(o, s.n), started = parseD(s.from) <= nowD, ended = parseD(s.to) <= nowD;
    const amt = b ? b.total : o.lines.reduce((t,l)=>t+priceLine(l, s.from, s.to, s.len).total, 0);
    const st = b ? `<span class="chip available">Billed ${fmtDate(b.at)}</span>` : back ? `<span class="chip neutral">On final invoice</span>` : ended ? `<span class="chip overdue">Due</span>` : started ? `<span class="chip rented">In progress</span>` : `<span class="chip neutral">Upcoming</span>`;
    const btn = b ? `<button class="btn sm" data-act="inv-open" data-no="${o.no}" data-n="${s.n}">${ico("printer")} INV-${o.no}-C${s.n}</button>` : (!back && started) ? `<button class="btn sm ${ended?"primary":""}" data-act="inv-open" data-no="${o.no}" data-n="${s.n}">Create invoice</button>` : "";
    return `<div class="line cyc-line"><div><b>Cycle ${s.n}</b> <span class="small muted num">${fmtDate(s.from)} – ${fmtDate(s.to)}${s.full?"":" · partial"}</span><div class="calc">${st}</div></div><div class="num">${b?"":'<span class="small muted">rental </span>'}${money(amt)}</div><span>${btn}</span></div>`; }).join("");
  return `<div class="subhead"><span>Cycle billing · ${esc(cycleName(o))}</span><span class="small num" style="text-transform:none;letter-spacing:0">${money(billed)} billed so far</span></div>
    <div class="lines">${rows || '<div class="muted small">No cycles yet.</div>'}</div>
    <span class="hint">${back ? "Anything not on a cycle invoice, plus return charges, is on the final invoice." : "A cycle shows Due once it ends. You can also bill the cycle in progress early. Delivery, sold items and charges go on the next invoice you create."}</span>`;
}
function markCycleBilled(no, n){
  const o = S.orders.find(x=>x.no===no); if (!o || billFor(o, n)) return;
  const d = invoiceData(o, {cycle:n});
  (o.bills ||= []).push({n, from:d.from, to:d.through, at:toLocalInput(now()), by:myName(), keys:unbilledKeys(o), total:+d.total.toFixed(2)});
  o.bills.sort((a,b)=>a.n-b.n);
  (o.log ||= []).push({by:myName(), at:toLocalInput(now()), text:`Cycle ${n} invoiced (INV-${o.no}-C${n}, ${money(d.total)})`});
  save(); toast(`Cycle ${n} marked billed`); openInvoice(no, n);
}
function voidCycleBill(no, n){
  const o = S.orders.find(x=>x.no===no), b = o && billFor(o, n); if (!b) return;
  o.bills = o.bills.filter(x=>x!==b); if (!o.bills.length) delete o.bills;
  (o.log ||= []).push({by:myName(), at:toLocalInput(now()), text:`Cycle ${n} invoice voided (INV-${o.no}-C${n})`});
  save(); toast(`INV-${o.no}-C${n} voided`); viewOrder(no);
}

function orderRentalEnd(o){ return o.status==="Returned" && o.returnedAt ? o.returnedAt : o.ret; }
function orderTotals(o){
  const end = orderRentalEnd(o);
  const lines = o.lines.map(l => ({...l, ...lineRental(o, l, end)}));
  const rental = lines.reduce((s,l)=>s+l.total,0);
  const delivery = (+o.delivery?.rate||0) * (+o.delivery?.trips||0);
  const charges = (o.charges||[]).reduce((s,c)=>s+(+c.qty||0)*(+c.rate||0),0);
  const bundleDisc = o.bundle ? lines.filter(l=>o.bundle.itemIds.includes(l.itemId)).reduce((s,l)=>s+l.total,0)*(+o.bundle.pct||0)/100 : 0;
  const sales = salesTotal(o);
  const total = rental + delivery + charges + sales + (+o.adjustment||0) - bundleDisc;
  return {lines,rental,delivery,charges,sales,total,bundleDisc};
}
/* inventory sold on a rental order: stock comes off when the order is out or returned, goes back if it's cancelled or the line is removed */
const canSell = i => !!i && (i.sellable ?? (+i.price > 0));
function salesTotal(o){ return (o.sales||[]).reduce((s,x)=>s+(+x.qty||0)*(+x.price||0),0); }
function syncSales(o, why){
  const want = {};
  if (["Rented Out","Returned"].includes(o.status)) (o.sales||[]).forEach(x=>{ if (x.invId && +x.qty>0) want[x.invId] = +((want[x.invId]||0) + +x.qty).toFixed(2); });
  const had = o.salesTaken||{}, at = toLocalInput(now()), by = myName();
  new Set([...Object.keys(want),...Object.keys(had)]).forEach(id=>{ const d = +((want[id]||0)-(had[id]||0)).toFixed(2); const it = invById(id); if (!d || !it) return;
    it.qty = +(it.qty - d).toFixed(2);
    (it.log ||= []).push({by,at,delta:-d,reason: d>0 ? `Sold on order #${o.no}${custById(o.customerId)?` (${custById(o.customerId).name})`:""}` : `Back in stock from order #${o.no}${why?` (${why})`:""}`}); });
  o.salesTaken = want;
}

/* ---------------- Status logic ---------------- */
const ACTIVE = ["Booked","Rented Out"];
const isReady = o => o.status==="Booked" && parseD(o.out) <= now();
const isOverdue = o => o.status==="Rented Out" && parseD(o.ret) < now();
function orderStatus(o){ return isOverdue(o) ? "Overdue" : isReady(o) ? "Ready to go out" : o.status; }
const SEVERITY = [["Low","Minor – Asset can keep working, fix when convenient"],["Medium","Needs repair soon – Limit use until fixed"],["High","Down or unsafe – Do not use until repaired"]];
const SEV_RANK = {High:0, Medium:1, Low:2};
function sevChip(v){ return v ? `<span class="chip sev-${v.toLowerCase()}">${esc(v)}</span>` : '<span class="muted">—</span>'; }
function chipFor(st){
  const cls = {...SR_CHIPS, "Rented Out":"rented","Available":"available","Booked":"reserved","Customer Pickup":"neutral","Customer Returned":"available","BHESC Delivered":"rented","BHESC Pickup":"checked","Ready to go out":"maint","Overdue":"overdue","In Maintenance":"maint","Returned":"returned","Retired":"retired","Cancelled":"cancelled","Open":"maint","Completed":"available","Checked Out":"checked","Queued":"neutral","In progress":"rented","Waiting on parts":"maint","Done":"available","In stock":"available","Low stock":"maint","Out of stock":"overdue","Admin":"checked","Manager":"rented","Rental Coordinator":"available","Mechanic":"maint","Field Crew":"neutral","Viewer":"neutral","Active":"available","Disabled":"cancelled","No password":"maint","Temporary password":"maint"}[st]||"neutral";
  return `<span class="chip ${cls}">${esc(st)}</span>`;
}
function itemById(id){ return S.items.find(i=>i.id===id); }
function custById(id){ return S.customers.find(c=>c.id===id); }
function locById(id){ return S.locations.find(l=>l.id===id); }
function yardId(){ return (S.locations.find(l=>l.isYard)||S.locations[0])?.id||""; }
function homeLoc(it){ return locById(it.locationId) || locById(yardId()); }
function homeLabel(it){ const l = homeLoc(it); return (l ? locLabel(l) : "Yard") + (it.home ? ` · ${it.home}` : ""); }
/* groups can have one level of sub-groups (g.parent = parent group's name) */
const byNum = (a,b) => a.localeCompare(b,undefined,{numeric:true});
function groupOf(name){ return S.groups.find(g=>g.name===name); }
function subGroups(name){ return S.groups.filter(g=>g.parent===name).sort((a,b)=>byNum(a.name,b.name)); }
function groupFamily(name){ return [name, ...subGroups(name).map(g=>g.name)]; }
function groupLabel(name){ const g = groupOf(name); return g?.parent ? `${g.parent} › ${name}` : name; }
function orderedGroups(){ const tops = S.groups.filter(g=>!g.parent || !groupOf(g.parent)).sort((a,b)=>byNum(a.name,b.name)); return tops.flatMap(g=>[g, ...subGroups(g.name)]); }
function groupOptions(selected, blank){ return (blank?`<option value="">${blank}</option>`:"") + orderedGroups().map(g=>`<option value="${esc(g.name)}" ${g.name===selected?"selected":""}>${g.parent&&groupOf(g.parent)?"\u00a0\u00a0\u00a0› ":""}${esc(g.name)}</option>`).join(""); }
/* A–Z (numbers in order) for drop-downs */
const az = (a,b) => String(a||"").localeCompare(String(b||""),undefined,{numeric:true,sensitivity:"base"});
const byName = list => list.slice().sort((a,b)=>az(a.name,b.name));
/* locations can sit under a parent location (one level) */
/* locations nest any number of levels: job › building › floor › room */
const locPath = l => { const out = []; let p = l; const seen = new Set(); while (p && p.parent && !seen.has(p.id)){ seen.add(p.id); p = S.locations.find(x=>x.id===p.parent && x.id!==p.id); if (!p || seen.has(p.id)) break; out.unshift(p); } return out; };
const locParent = l => { const pth = l ? locPath(l) : []; return pth.length ? pth[pth.length-1] : null; };
const locDepth = l => locPath(l).length;
const locKids = id => S.locations.filter(l=>l.parent===id && l.id!==id && locParent(l)?.id===id).sort((a,b)=>az(locLabel(a),locLabel(b)));
const locFamily = id => [id, ...locKids(id).flatMap(l=>locFamily(l.id))];
const locTree = l => [l, ...locKids(l.id).flatMap(locTree)];
const sortedLocs = () => S.locations.filter(l=>!locParent(l)).sort((a,b)=>az(locLabel(a),locLabel(b))).flatMap(locTree);
const locOpt = l => "\u00a0\u00a0\u00a0".repeat(locDepth(l)) + (locDepth(l) ? "› " : "") + locLabel(l);
/* inventory lives in the Equipment Shop by default (falls back to the yard if there's no shop location) */
function invHomeId(locs){ locs = locs || S.locations; return (locs.find(l=>/equipment shop/i.test(l.name)) || locs.find(l=>/\bshop\b/i.test(l.name) && !l.isYard) || locs.find(l=>l.isYard) || {}).id; }
function keptAt(x){ const l = locById(x.locationId); if (!l) return x.location || "—"; const up = locPath(l).slice(1); return [...up.map(a=>a.name), locLabel(l)].join(" › "); }
function keptAtSelect(id, cur){ return `<select id="${id}" name="locationId">${sortedLocs().map(l=>`<option value="${l.id}" ${l.id===(cur||yardId())?"selected":""}>${esc(locOpt(l))}${l.isYard?" (default)":""}</option>`).join("")}</select>`; }
function locContact(l){ if (!l) return ""; const c = l.contactId ? custById(l.contactId) : null; return c ? c.name : (l.contact||""); }
function locLabel(l){ return l ? (l.code ? `${l.code} - ${l.name}` : l.name) : "—"; }
function activeOrderFor(itemId){ return S.orders.find(o=>o.status==="Rented Out" && o.lines.some(l=>l.itemId===itemId)); }
function nextReservation(itemId){ return S.orders.filter(o=>o.status==="Booked" && o.lines.some(l=>l.itemId===itemId)).sort((a,b)=>parseD(a.out)-parseD(b.out))[0]; }
function openSR(itemId){ return S.maintenance.find(m=>m.itemId===itemId && m.status==="Open"); }
/* an open SR takes the asset out of service once it's deadlined or accepted into the shop; an operational problem that's only reported stays rentable */
function openMaint(itemId){ return S.maintenance.find(m=>m.itemId===itemId && m.status==="Open" && (m.opStatus==="Deadlined" || (m.stage && m.stage!=="Submitted"))); }
function openCustody(itemId){ return S.custody.find(c=>c.kind==="asset" && c.itemId===itemId && !c.inAt); }
function stockOut(stockId){ return S.custody.filter(c=>c.kind==="stock" && c.stockId===stockId && !c.inAt).reduce((s,c)=>s+(+c.qty||0),0); }
const custodyOverdue = c => !c.inAt && c.dueAt && parseD(c.dueAt) < now();
function stockById(id){ return S.stock.find(s=>s.id===id); }
function custodyLabel(c){ if (c.kind==="asset"){ const it=itemById(c.itemId); return `#${c.itemId} ${it?it.name:"(deleted)"}`; } if (c.kind==="inv"){ const i=S.inventory.find(x=>x.id===c.invId); return i?i.name:"(deleted inventory)"; } const s=stockById(c.stockId); return s?s.name:"(deleted stock)"; }
function itemStatus(it){
  if (it.status==="Retired") return "Retired";
  if (openMaint(it.id)) return "In Maintenance";
  if (openCustody(it.id)) return "Checked Out";
  const o = activeOrderFor(it.id);
  if (o) return isOverdue(o) ? "Overdue" : "Rented Out";
  return "Available";
}
function conflicts(itemId, out, ret, exceptNo){
  const a = parseD(out), b = parseD(ret);
  return S.orders.filter(o => o.no!==exceptNo && ACTIVE.includes(o.status) && o.lines.some(l=>l.itemId===itemId)).filter(o=>{
    const s = parseD(o.out), e = isOverdue(o) ? new Date(Math.max(now(), parseD(o.ret))) : parseD(o.ret);
    return s < b && a < e;
  });
}
function utilization(itemId, days=90){
  const end = now(), start = new Date(end - days*DAY); let used = 0;
  for (const o of S.orders){
    if (!["Rented Out","Returned"].includes(o.status) || !o.lines.some(l=>l.itemId===itemId)) continue;
    const s = parseD(o.out), e = o.status==="Returned" ? parseD(o.returnedAt||o.ret) : end;
    const ov = Math.min(e,end) - Math.max(s,start); if (ov>0) used += ov;
  }
  return Math.min(1, used / (days*DAY));
}

/* ---------------- Routing & UI state ---------------- */
const ui = { orders:{tab:"active",q:"",status:"",sort:"ret",dir:1,sel:new Set(),filter:""}, items:{tab:"all",q:"",cat:"",exact:""}, groups:{open:null}, locs:{open:null}, maint:{tab:"open",sev:"",type:"",q:""}, reports:{tab:"revenue",period:"ytd"} };
function route(){ const h = location.hash.replace("#","");
  if (h==="available"){ if (isSuper()) ui.myrep.tab = "avail"; else ui.reports.tab = "avail"; history.replaceState(null,"","#"+(isSuper()?"my-reports":"reports")); return isSuper()?"my-reports":"reports"; }
  if (isSuper()) return SUPER_ROUTES.includes(h) ? h : "request";
  if (h==="asset-optimizer"||h==="workshop"||h==="bundles") return "dashboard";
  return (NAV.some(n=>n[0]===h) || ITEM_KEYS.includes(h) || h==="users" || h==="requests" || /^asset(-[A-Za-z0-9-]+)?$/.test(h)) ? h : "dashboard"; }
function navList(){ if (isSuper()) return [["request","Request"],["my-requests","My requests"],["my-reports","My reports"]]; const n = NAV.slice(); n.splice(2,0,["requests","Requests"]); return n; }
function go(to){ if (location.hash === "#"+to) render(); else location.hash = to; }
document.addEventListener("click", e=>{ const a = e.target.closest('a[href="#locations"]'); if (a && ui.locs.open){ ui.locs.open = null; if (route()==="locations") rerender(); } });
document.addEventListener("click", e=>{ const a = e.target.closest('a[href="#groups"]'); if (a && ui.groups.open){ ui.groups.open = null; if (route()==="groups") rerender(); } });
window.addEventListener("hashchange", ()=>{ closeModal(); draft=null; ci=null; codraft=null; coin=null; mdraft=null; render(); });

function renderRail(){
  const cur = route();
  const u = me();
  $("#rail").innerHTML = `<button class="avatar-btn" id="avatar-btn" data-act="profile-menu" aria-haspopup="menu" title="${esc(u?u.name+" · "+u.role:"")}">${avatarHtml(u)}</button>` +
    navList().filter(([k])=>!PAGE_PERM[k]||can(PAGE_PERM[k])).map(([k,l])=>`${k==="setup"?'<div class="spacer"></div>':''}<a href="#${k}" ${k==="items"?'id="items-link" aria-haspopup="true"':''} ${(k===cur||(k==="items"&&(ITEM_KEYS.includes(cur)||/^asset(-|$)/.test(cur))))?'aria-current="page"':''}>${ico(k)}<span>${l}</span>${k==="requests"&&pendingRequests().length?`<span class="rail-badge">${pendingRequests().length}</span>`:""}</a>`).join("");
  const link = $("#items-link"); if (!link) return;
  link.addEventListener("mouseenter", showFlyout); link.addEventListener("mouseleave", ()=>hideFlyoutSoon());
  link.addEventListener("focus", showFlyout);
}
let flyTimer;
function showFlyout(){
  if (matchMedia("(max-width:760px)").matches) return;
  clearTimeout(flyTimer); let f = $("#flyout");
  if (!f){ f = document.createElement("nav"); f.id="flyout"; f.className="flyout"; f.setAttribute("aria-label","Items");
    f.addEventListener("mouseenter",()=>clearTimeout(flyTimer)); f.addEventListener("mouseleave",()=>hideFlyoutSoon()); f.addEventListener("click",()=>f.remove());
    f.addEventListener("focusout",e=>{ if(!f.contains(e.relatedTarget)) hideFlyoutSoon(); });
    document.body.appendChild(f); }
  const cur = route();
  f.innerHTML = ITEM_SUB.map(([k,l,i,isNew])=>`<a href="#${k}" ${k===cur?'aria-current="page"':''}>${ico(i)}<span>${l}</span>${isNew?'<span class="new">NEW</span>':""}</a>`).join("");
  const r = $("#items-link").getBoundingClientRect(); f.style.left = r.right+"px"; f.style.top = Math.max(8, r.top)+"px";
}
function hideFlyoutSoon(){ clearTimeout(flyTimer); flyTimer = setTimeout(()=>$("#flyout")?.remove(), 220); }
function subnav(){ const cur = route(); return `<nav class="subnav" aria-label="Items sections">${ITEM_SUB.map(([k,l])=>`<a href="#${k}" ${k===cur?'aria-current="page"':''}>${l}</a>`).join("")}</nav>`; }
function initials(n){ return (n||"").split(/\s+/).map(x=>x[0]).join("").slice(0,2).toUpperCase(); }
function notifItems(){
  const out = [], add = (n, tone, icon, title, sub, attrs) => { if (n) out.push({n, tone, icon, title, sub, attrs}); };
  if (can("admin")){ const rq = pendingRequests(); add(rq.length, "bad", "requests", `${rq.length} equipment request${rq.length>1?"s":""} waiting`, rq.slice(0,3).map(r=>`${rqNo(r)} ${r.byName}`).join(" · "), 'data-act="goto" data-to="requests"'); }
  const od = S.orders.filter(isOverdue).sort((a,b)=>parseD(a.ret)-parseD(b.ret));
  add(od.length, "bad", "orders", `${od.length} overdue return${od.length>1?"s":""}`, od.slice(0,3).map(o=>`#${o.no} ${custById(o.customerId)?.name||""} · due ${fmtDate(o.ret)}`).join(" · "), 'data-act="goto" data-to="orders" data-filter="overdue"');
  const cl = S.custody.filter(c=>!c.inAt && custodyOverdue(c));
  add(cl.length, "bad", "checkout", `${cl.length} crew check-out${cl.length>1?"s":""} past due`, cl.slice(0,3).map(c=>`${custodyLabel(c)} · ${c.holder}`).join(" · "), 'data-act="goto" data-to="checkout" data-cotab="overdue"');
  const cb = S.orders.filter(o=>cycleDue(o).length);
  add(cb.length, "warn", "orders", `${cb.length} cycle bill${cb.length>1?"s":""} due`, cb.slice(0,3).map(o=>`#${o.no} ${custById(o.customerId)?.name||""} · cycle ${cycleDue(o)[0].n}`).join(" · "), 'data-act="goto" data-to="orders" data-filter="billdue"');
  const rd = S.orders.filter(o=>o.status==="Booked" && isReady(o));
  add(rd.length, "info", "orders", `${rd.length} booked order${rd.length>1?"s":""} ready to go out`, rd.slice(0,3).map(o=>`#${o.no} · out ${fmtDate(o.out)}`).join(" · "), 'data-act="goto" data-to="orders" data-filter="ready"');
  const dl = S.maintenance.filter(m=>m.status==="Open" && m.opStatus==="Deadlined"), qc = S.maintenance.filter(m=>m.status==="Open" && m.stage==="QC inspection"), sub = S.maintenance.filter(m=>m.status==="Open" && m.stage==="Submitted");
  add(dl.length, "warn", "maintenance", `${dl.length} deadlined asset${dl.length>1?"s":""}`, dl.slice(0,3).map(m=>`${srNo(m)} #${m.itemId}`).join(" · "), 'data-act="goto" data-to="maintenance" data-mtab="deadline"');
  add(sub.length, "info", "maintenance", `${sub.length} service request${sub.length>1?"s":""} to accept`, sub.slice(0,3).map(m=>`${srNo(m)} #${m.itemId}`).join(" · "), 'data-act="goto" data-to="maintenance" data-mtab="submitted"');
  add(qc.length, "info", "maintenance", `${qc.length} ready for final inspection`, qc.slice(0,3).map(m=>`${srNo(m)} #${m.itemId}`).join(" · "), 'data-act="goto" data-to="maintenance" data-mtab="qc"');
  const low = S.inventory.filter(i=>!i.retired && invStatus(i)!=="In stock");
  add(low.length, "warn", "inventory", `${low.length} inventory item${low.length>1?"s":""} low or out`, low.slice(0,3).map(i=>`${i.name} (${i.qty} ${i.unit||""})`).join(" · "), 'data-act="goto" data-to="inventory"');
  return out;
}
function bellCount(){ return (can("admin") ? pendingRequests().length : 0) + S.orders.filter(isOverdue).length + S.orders.filter(o=>cycleDue(o).length).length + S.custody.filter(c=>!c.inAt && custodyOverdue(c)).length; }
function updateBell(){ const n = bellCount(); const b = $("#bellcount"); b.hidden = !n; b.textContent = n;
  const btn = $("#bell-btn"); if (btn) btn.title = n ? `${n} item${n>1?"s":""} need attention` : "Notifications";
  document.body.classList.toggle("role-super", isSuper()); }
function toggleNotif(){
  let m = $("#notif"); if (m) { m.remove(); return; }
  const items = notifItems();
  m = document.createElement("div"); m.id = "notif"; m.className = "pmenu notif"; m.setAttribute("role","menu");
  m.innerHTML = `<div class="nt-head"><b>Notifications</b><span class="small muted">${items.length ? `${items.length} thing${items.length>1?"s":""} to look at` : ""}</span></div>
    ${items.length ? items.map(x=>`<button role="menuitem" class="nt-item" ${x.attrs}><span class="nt-ic ${x.tone}">${ico(x.icon)}</span><span class="nt-txt"><b>${esc(x.title)}</b>${x.sub?`<small>${esc(x.sub)}</small>`:""}</span><span class="m-chev">›</span></button>`).join("")
      : `<div class="nt-empty"><span class="nt-ok">✓</span><b>You're all caught up</b><span class="small muted">No overdue returns, requests or past-due check-outs.</span></div>`}`;
  document.body.appendChild(m);
  const r = $("#bell-btn").getBoundingClientRect();
  m.style.top = (r.bottom + 8) + "px"; m.style.left = Math.max(8, Math.min(window.innerWidth - 348, r.right - 340)) + "px";
}

function band(title, sub, actions=""){ return `${ITEM_KEYS.includes(route())?subnav():`<div class="crumb">${esc(title)}</div>`}<div class="band"><h1>${esc(title)}</h1>${actions?`<div class="actions">${actions}</div>`:""}${sub?`<div class="sub">${sub}</div>`:""}</div>`; }
function tabs(list, cur, act){ return `<div class="tabs" role="tablist">${list.map(([k,l,c])=>`<button role="tab" aria-selected="${k===cur}" data-act="${act}" data-v="${k}">${l}${c!=null?`<span class="count">${c}</span>`:""}</button>`).join("")}</div>`; }

function guardedView(r){ if (PAGE_PERM[r] && !can(PAGE_PERM[r])) return band("No access","") + `<div class="content"><div class="empty"><strong>Your role (${esc(me().role)}) can't open this page</strong>Ask an administrator if you need access.</div></div>`; if (VIEWS[r]) return VIEWS[r](); return vAsset(assetIdFromRoute(r)); }
function brandHtml(cls){ const s = S.settings; return s.logo ? `<img class="brand-img ${cls||""}" src="${s.logo}" alt="${esc(s.company||"Company")} logo">` : `BH<small>EQUIPMENT</small>`; }
function updateBrand(){ const el = $("#top-logo"); if (el){ el.innerHTML = brandHtml("top"); el.classList.toggle("has-img", !!S.settings.logo); } }
function render(){
  updateBrand();
  $("#pmenu")?.remove();
  if (!authGate()) return;
  renderRail(); updateBell(); applyPerms($(".topbar"));
  const r = route();
  $("#view").innerHTML = guardedView(r); applyPerms($("#view")); cardify($("#view")); initResizable($("#view")); renderMobileChrome();
  window.scrollTo(0,0);
}
function rerender(){ updateBrand(); const y = window.scrollY; const f = document.activeElement?.id; const pos = document.activeElement?.selectionStart;
  renderRail(); updateBell();
  if (!me()) { render(); return; }
  const r = route(); $("#view").innerHTML = guardedView(r); applyPerms($("#view")); cardify($("#view")); initResizable($("#view")); renderMobileChrome();
  window.scrollTo(0,y); if (f){ const el = document.getElementById(f); if (el){ el.focus(); try{ if(pos!=null) el.setSelectionRange(pos,pos);}catch(e){} } } }

/* ---------------- Dashboard ---------------- */
/* ---------------- Dashboard: widgets each user can move, resize, remove and add back ---------------- */
const DASH_DEFAULT = [["k-fleet","s"],["k-onrent","s"],["k-avail","s"],["k-checked","s"],["k-maint","s"],["k-overdue","s"],["k-runrate","s"],["k-billed","s"],
  ["p-overdue","m"],["p-soon","m"],["p-cats","m"],["p-booked","m"]];
ui.dash = {edit:false, snap:null};
/* free-placement grid: 4 columns, fixed row height; each widget has a column (x), row (y) and a size */
const DCOLS = 4, DROW = 112, DGAP = 14;
const isStatW = id => id.startsWith("k-");
function dwSpan(w){ const wd = {s:1, m:2, l:DCOLS}[w.size||"m"] || 2; return {w:wd, h:isStatW(w.id) ? 1 : 3}; }
const dwHit = (a, b) => a.x < b.x + dwSpan(b).w && b.x < a.x + dwSpan(a).w && a.y < b.y + dwSpan(b).h && b.y < a.y + dwSpan(a).h;
function dwSpot(placed, w){ const sp = dwSpan(w); for (let y = 0; y < 400; y++) for (let x = 0; x <= DCOLS - sp.w; x++){ const t = {...w, x, y}; if (!placed.some(p=>dwHit(t,p))) return {x, y}; } return {x:0, y:0}; }
function dashLayout(){
  const u = me(), src = u && Array.isArray(u.dash) ? u.dash : DASH_DEFAULT.map(([id,size])=>({id,size}));
  const placed = [];
  src.filter(w=>DASH_WIDGETS[w.id]).forEach(w0=>{ const w = {...w0}, sp = dwSpan(w);
    if (w.x==null || w.y==null){ Object.assign(w, dwSpot(placed, w)); }
    w.x = Math.max(0, Math.min(w.x, DCOLS - sp.w)); w.y = Math.max(0, w.y);
    while (placed.some(p=>dwHit(w,p))) w.y++;
    placed.push(w); });
  return placed;
}
/* push anything the moved widget now overlaps further down, cascading */
function dwResolve(l, i){
  const fixed = [l[i]], rest = l.filter((_,k)=>k!==i).sort((a,b)=>a.y-b.y || a.x-b.x);
  rest.forEach(o=>{ let hit; while ((hit = fixed.filter(f=>dwHit(o,f))).length) o.y = Math.max(...hit.map(f=>f.y + dwSpan(f).h)); fixed.push(o); });
}
function setDashLayout(l){ const u = me(); if (!u) return; u.dash = l.map(({id,size,x,y})=>({id,size,x,y})); save(); rerender(); }
let DASH_WIDGETS = {};
function dashData(){
  const items = S.items.filter(i=>i.status!=="Retired" && !i.shopOnly), st = items.map(itemStatus), d = now();
  const active = S.orders.filter(o=>o.status==="Rented Out"), overdue = S.orders.filter(isOverdue);
  const mStart = new Date(d.getFullYear(), d.getMonth(), 1);
  return {items, st, d, active, overdue,
    onRent: st.filter(s=>s==="Rented Out"||s==="Overdue").length, avail: st.filter(s=>s==="Available").length, maint: st.filter(s=>s==="In Maintenance").length,
    checked: st.filter(s=>s==="Checked Out").length,
    monthlyRun: active.reduce((s,o)=>s+o.lines.reduce((a,l)=>{const p=priceLine(l,"2000-01-01T00:00",new Date(new Date("2000-01-01T00:00").getTime()+(+S.settings.cycle||28)*DAY).toISOString()); return a+p.total;},0),0),
    billedMonth: S.orders.filter(o=>o.status!=="Cancelled" && parseD(o.out)>=mStart).reduce((s,o)=>s+orderTotals(o).total,0),
    soon: active.filter(o=>!isOverdue(o) && parseD(o.ret) - d < 7*DAY).sort((a,b)=>parseD(a.ret)-parseD(b.ret)),
    booked: S.orders.filter(o=>o.status==="Booked").sort((a,b)=>parseD(a.out)-parseD(b.out))};
}
function dashOrderRow(o, d){ const c = custById(o.customerId), l = locById(o.locationId);
  return `<div class="list-row"><div><button class="linkbtn" data-act="view-order" data-no="${o.no}">#${o.no}</button> <span class="muted">·</span> ${esc(o.lines.map(x=>"#"+x.itemId).join(", "))}<div class="small muted">${esc(locLabel(l))}${c?` · ${esc(c.name)}`:""}</div></div><div style="text-align:right">${chipFor(orderStatus(o))}<div class="small muted num">${isOverdue(o)?Math.floor((d-parseD(o.ret))/DAY)+" days late":"Due "+fmtDate(o.ret)}</div></div></div>`; }
const kpiW = (lbl, val, note, bad) => ({stat:true, html:`<div class="kpi ${bad?"bad":""}"><span class="lbl">${lbl}</span><span class="val num">${val}</span><span class="note">${note}</span></div>`});
const panelW = (title, body, link) => ({html:`<div class="panel"><h3><span>${title}</span>${link||""}</h3><div class="body">${body}</div></div>`});
DASH_WIDGETS = {
  "k-fleet":   {name:"Fleet size", r:D=>{ const shop = S.items.filter(i=>i.shopOnly && i.status!=="Retired").length; return kpiW("Rental fleet", D.items.length, shop ? `plus ${shop} shop-use asset${shop>1?"s":""}` : "assets in inventory"); }},
  "k-onrent":  {name:"On rent", r:D=>kpiW("On rent", D.onRent, `${Math.round(D.onRent/(D.items.length||1)*100)}% of fleet out`)},
  "k-avail":   {name:"Available", r:D=>kpiW("Available", D.avail, "ready in the yard")},
  "k-checked": {name:"Checked out", r:D=>kpiW("Checked out", D.checked, "with your crews, not billed")},
  "k-maint":   {name:"In maintenance", r:D=>kpiW("In maintenance", D.maint, "open service tickets")},
  "k-overdue": {name:"Overdue returns count", r:D=>kpiW("Overdue returns", D.overdue.length, "past return date, still out", D.overdue.length)},
  "k-runrate": {name:"Monthly run rate", r:D=>kpiW("Monthly run rate", money0(D.monthlyRun), `equipment on rent, per ${S.settings.cycle}-day cycle`)},
  "k-billed":  {name:"Billed this month", r:D=>kpiW("Billed this month", money0(D.billedMonth), "orders that went out this month")},
  "k-booked":  {name:"Booked orders count", r:D=>kpiW("Booked", D.booked.length, "orders holding equipment")},
  "k-lowinv":  {name:"Low inventory count", r:D=>{ const low = S.inventory.filter(i=>!i.retired && invStatus(i)!=="In stock").length; return kpiW("Low or out of stock", low, "inventory items to reorder", low); }},
  "p-overdue": {name:"Overdue returns list", r:D=>panelW("Overdue returns", `<div class="list">${D.overdue.length?D.overdue.slice().sort((a,b)=>parseD(a.ret)-parseD(b.ret)).slice(0,8).map(o=>dashOrderRow(o,D.d)).join(""):'<div class="muted">Nothing overdue.</div>'}</div>`, `<button class="btn sm" data-act="goto" data-to="orders" data-filter="overdue">View all</button>`)},
  "p-soon":    {name:"Due back in 7 days", r:D=>panelW("Due back in the next 7 days", `<div class="list">${D.soon.length?D.soon.map(o=>dashOrderRow(o,D.d)).join(""):'<div class="muted">No returns due this week.</div>'}</div>`)},
  "p-cats":    {name:"Fleet out by group", r:D=>{ const cats = {}; D.items.forEach((it,i)=>{ const c = cats[it.category] ||= {n:0,out:0}; c.n++; if (D.st[i]==="Rented Out"||D.st[i]==="Overdue") c.out++; });
    return panelW("Fleet out by group", `<div class="bars">${Object.entries(cats).sort((a,b)=>b[1].out/b[1].n - a[1].out/a[1].n).map(([k,v])=>`<div class="bar-row"><span>${esc(k)}</span><div class="track"><div class="fill" style="width:${v.out/v.n*100}%"></div></div><span class="v">${v.out} / ${v.n}</span></div>`).join("")}</div>`); }},
  "p-booked":  {name:"Booked orders list", r:D=>panelW("Booked orders", `<div class="list">${D.booked.length?D.booked.slice(0,8).map(o=>`<div class="list-row"><div><button class="linkbtn" data-act="view-order" data-no="${o.no}">#${o.no}</button> ${esc(o.lines.map(x=>"#"+x.itemId).join(", "))}<div class="small muted">${esc(locLabel(locById(o.locationId)))}</div></div><div class="small num">Out ${fmtDate(o.out)}</div></div>`).join(""):'<div class="muted">Nothing booked. Choose <b>Booked</b> on a new order to hold equipment for future dates.</div>'}</div>`)},
  "p-maint":   {name:"Open work orders", r:D=>{ const open = S.maintenance.filter(m=>m.status==="Open").sort((a,b)=>(SEV_RANK[a.severity]??3)-(SEV_RANK[b.severity]??3) || b.date.localeCompare(a.date));
    return panelW("Open work orders", `<div class="list">${open.length?open.slice(0,8).map(m=>`<div class="list-row"><div><button class="linkbtn" data-act="view-item" data-id="${esc(m.itemId)}">#${esc(m.itemId)}</button> ${esc(itemById(m.itemId)?.name||"")}<div class="small muted">${esc(m.desc||m.type)}</div></div><div style="text-align:right">${sevChip(m.severity)}<div class="small muted">${esc(m.stage||"Queued")}</div></div></div>`).join(""):'<div class="muted">No open work orders.</div>'}</div>`, `<button class="btn sm" data-act="goto" data-to="maintenance">View all</button>`); }},
  "p-crews":   {name:"Out with crews", r:D=>{ const out = S.custody.filter(c=>!c.inAt).sort((a,b)=>(a.dueAt||"9").localeCompare(b.dueAt||"9"));
    return panelW("Out with crews", `<div class="list">${out.length?out.slice(0,8).map(c=>`<div class="list-row"><div><b>${esc(custodyLabel(c))}</b>${c.qty>1?` × ${c.qty}`:""}<div class="small muted">${esc(c.holder)} · ${esc(locLabel(locById(c.locationId)))}</div></div><div class="small num" ${custodyOverdue(c)?'style="color:var(--bad)"':""}>${c.dueAt?(custodyOverdue(c)?"Late · ":"Due ")+fmtDate(c.dueAt):"No due date"}</div></div>`).join(""):'<div class="muted">Nothing is checked out to crews.</div>'}</div>`, `<button class="btn sm" data-act="goto" data-to="checkout">View all</button>`); }},
  "p-lowinv":  {name:"Inventory to reorder", r:D=>{ const low = S.inventory.filter(i=>!i.retired && invStatus(i)!=="In stock");
    return panelW("Inventory to reorder", `<div class="list">${low.length?low.map(i=>`<div class="list-row"><div><b>${esc(i.name)}</b><div class="small muted">${esc(keptAt(i))}</div></div><div style="text-align:right"><span class="chip ${i.qty<=0?"overdue":"maint"}">${invStatus(i)}</span><div class="small muted num">${i.qty} ${esc(i.unit||"")} on hand</div></div></div>`).join(""):'<div class="muted">Everything is above its reorder point.</div>'}</div>`, `<button class="btn sm" data-act="goto" data-to="inventory">View all</button>`); }},
  "p-recent":  {name:"Recent orders", r:D=>panelW("Recent orders", `<div class="list">${S.orders.slice().sort((a,b)=>b.no-a.no).slice(0,8).map(o=>dashOrderRow(o,D.d)).join("")||'<div class="muted">No orders yet.</div>'}</div>`, `<button class="btn sm" data-act="goto" data-to="orders" data-otab="all">View all</button>`)},
  "p-locs":    {name:"On rent by location", r:D=>{ const by = {}; D.active.forEach(o=>{ const k = o.locationId; (by[k] ||= {n:0,orders:0}); by[k].n += o.lines.length; by[k].orders++; });
    const rows = Object.entries(by).sort((a,b)=>b[1].n-a[1].n), max = Math.max(1,...rows.map(r=>r[1].n));
    return panelW("On rent by location", `<div class="bars">${rows.length?rows.slice(0,8).map(([id,v])=>`<div class="bar-row"><span>${esc(locLabel(locById(id)))}</span><div class="track"><div class="fill" style="width:${v.n/max*100}%"></div></div><span class="v">${v.n}</span></div>`).join(""):'<div class="muted">Nothing on rent.</div>'}</div>`, `<button class="btn sm" data-act="goto" data-to="reports">Report</button>`); }},
};
function vDashboard(){
  const D = dashData(), E = ui.dash.edit, layout = dashLayout(), shown = new Set(layout.map(w=>w.id));
  const hidden = Object.keys(DASH_WIDGETS).filter(id=>!shown.has(id));
  const tool = (w,i) => `<div class="dw-tools"><span class="dw-handle" title="Drag anywhere to move">⋮⋮</span><span class="dw-name">${esc(DASH_WIDGETS[w.id].name)}</span>
    <span class="dw-btns"><span class="dw-size">${[["s","S","Small"],["m","M","Half width"],["l","L","Full width"]].map(([k,t,l])=>`<button class="dw-b ${w.size===k?"on":""}" data-act="dw-size" data-i="${i}" data-v="${k}" title="${l}">${t}</button>`).join("")}</span>
    <button class="dw-b x" data-act="dw-remove" data-i="${i}" title="Remove from dashboard" aria-label="Remove">×</button></span></div>`;
  return band("Dashboard", `${esc(S.settings.company)} fleet as of ${fmtDate(toLocalInput(D.d))}`, E
      ? `<button class="btn" data-act="dw-reset">Reset to default</button><button class="btn primary" data-act="dw-edit">Done</button>`
      : `<button class="btn d-only" data-act="dw-edit">${ico("edit")} Customize</button><button class="btn primary" data-act="new-order">+ Add new order</button>`) + `
  <div class="content">
    ${mobileDashboard()}
    ${E?`<div class="dw-tray d-only"><div class="dw-tray-h"><b>Customize your dashboard</b><span class="small muted">Grab a widget and drop it anywhere on the grid. Gaps are fine, and widgets in the way move down. Pick a size (S, M, L) or × to remove. Your layout is saved to your profile.</span><span class="small dw-narrow" style="color:var(--warn)">Make the window wider to drag widgets around. On narrower screens they stack in order.</span></div>
      ${hidden.length?`<div class="dw-add"><span class="small muted">Add a widget:</span>${hidden.map(id=>`<button class="btn sm" data-act="dw-add" data-v="${id}">+ ${esc(DASH_WIDGETS[id].name)}</button>`).join("")}</div>`:`<div class="small muted">Every widget is on your dashboard.</div>`}</div>`:""}
    <div class="dash-grid ${E?"editing":""}" style="--drow:${DROW}px;--dgap:${DGAP}px">
      ${layout.map((w,i)=>{ const out = DASH_WIDGETS[w.id].r(D), sp = dwSpan(w);
        return `<div class="dw dw-${w.size||"m"} ${out.stat?"is-stat":"d-only"} ${ui.dash.snap===w.id?"snap":""}" data-i="${i}" data-id="${w.id}" data-x="${w.x}" data-y="${w.y}" data-w="${sp.w}" data-h="${sp.h}" style="grid-column:${w.x+1} / span ${sp.w};grid-row:${w.y+1} / span ${sp.h};order:${w.y*10+w.x};--w2:${sp.w>1?2:1};--h:${sp.h}">${E?tool(w,i):""}${out.html}</div>`; }).join("") || `<div class="empty" style="grid-column:1/-1"><strong>Your dashboard is empty</strong>Choose Customize to add widgets.</div>`}
      ${E?`<div class="dw-fill" style="grid-column:1;grid-row:${Math.max(0,...layout.map(w=>w.y+dwSpan(w).h))+4}"></div>`:""}
    </div>
  </div>`;
}
function dashEdit(fn){ const l = dashLayout().map(w=>({...w})); fn(l); setDashLayout(l); }
let dwDrag = null;
const dwWide = () => matchMedia("(min-width:1181px)").matches;
document.addEventListener("pointerdown", e=>{
  const w = e.target.closest?.(".dash-grid.editing .dw"); if (!w || e.button>0 || !dwWide()) return;
  if (e.target.closest("button, a, input, select, textarea")) return;
  e.preventDefault();
  const grid = w.parentElement, r = w.getBoundingClientRect();
  const ghost = w.cloneNode(true); ghost.classList.add("dw-ghost"); ghost.classList.remove("snap");
  Object.assign(ghost.style, {width:r.width+"px", height:r.height+"px", left:r.left+"px", top:r.top+"px", gridColumn:"", gridRow:""}); document.body.appendChild(ghost);
  const ph = document.createElement("div"); ph.className = "dw-drop"; ph.style.gridColumn = w.style.gridColumn; ph.style.gridRow = w.style.gridRow; grid.appendChild(ph);
  w.classList.add("dragging"); document.body.classList.add("dw-dragging");
  dwDrag = {i:+w.dataset.i, w, grid, ghost, ph, dx:e.clientX-r.left, dy:e.clientY-r.top, sw:+w.dataset.w, sh:+w.dataset.h, x:+w.dataset.x, y:+w.dataset.y, id:w.dataset.id};
});
document.addEventListener("pointermove", e=>{
  if (!dwDrag) return; const d = dwDrag, gx = e.clientX - d.dx, gy = e.clientY - d.dy;
  d.ghost.style.left = gx + "px"; d.ghost.style.top = gy + "px";
  const g = d.grid.getBoundingClientRect(), colW = (g.width - DGAP*(DCOLS-1)) / DCOLS;
  d.x = Math.max(0, Math.min(DCOLS - d.sw, Math.round((gx - g.left) / (colW + DGAP))));
  d.y = Math.max(0, Math.round((gy - g.top) / (DROW + DGAP)));
  d.ph.style.gridColumn = `${d.x+1} / span ${d.sw}`; d.ph.style.gridRow = `${d.y+1} / span ${d.sh}`;
  if (e.clientY > innerHeight - 50) scrollBy(0, 14); else if (e.clientY < 70) scrollBy(0, -14);
});
function dwEnd(cancel){
  if (!dwDrag) return; const d = dwDrag; dwDrag = null;
  d.ghost.remove(); d.ph.remove(); d.w.classList.remove("dragging"); document.body.classList.remove("dw-dragging");
  if (cancel) return;
  ui.dash.snap = d.id; dashEdit(l=>{ const k = l.findIndex(w=>w.id===d.id); l[k].x = d.x; l[k].y = d.y; dwResolve(l, k); });
  setTimeout(()=>{ if (ui.dash.snap===d.id) ui.dash.snap = null; }, 700);
}
document.addEventListener("pointerup", ()=>dwEnd(false));
document.addEventListener("pointercancel", ()=>dwEnd(true));
document.addEventListener("keydown", e=>{ if (e.key==="Escape" && dwDrag) dwEnd(true); });

/* ---------------- Orders ---------------- */
function vOrders(){
  const u = ui.orders, d = now();
  if (u.tab==="reserved") u.tab = "active";
  const counts = {active:S.orders.filter(o=>ACTIVE.includes(o.status)).length, returned:S.orders.filter(o=>o.status==="Returned"||o.status==="Cancelled").length, all:S.orders.length};
  let list = S.orders.filter(o => u.tab==="all" ? true : u.tab==="active" ? ACTIVE.includes(o.status) : o.status==="Returned"||o.status==="Cancelled");
  if (u.filter==="booked") list = list.filter(o=>o.status==="Booked");
  if (u.filter==="out") list = list.filter(o=>o.status==="Rented Out");
  if (u.filter==="overdue") list = list.filter(isOverdue);
  if (u.filter==="soon") list = list.filter(o=>o.status==="Rented Out" && !isOverdue(o) && parseD(o.ret)-d<7*DAY);
  if (u.filter==="ready") list = list.filter(isReady);
  if (u.filter==="pickup") list = list.filter(o=>o.status==="Rented Out"&&o.pickup==="BHESC Pickup");
  if (u.filter==="billdue") list = list.filter(o=>cycleDue(o).length);
  if (u.filter==="cycle") list = list.filter(o=>o.cycle);
  if (u.status) list = list.filter(o=>orderStatus(o)===u.status);
  if (u.q){ const q = u.q.toLowerCase(); list = list.filter(o => orderSearchText(o).includes(q)); }
  const key = {no:o=>o.no, cust:o=>custById(o.customerId)?.name||"", price:o=>orderTotals(o).total, out:o=>parseD(o.out), ret:o=>parseD(o.ret), status:o=>orderStatus(o)}[u.sort];
  list.sort((a,b)=>{ const x=key(a), y=key(b); return (x>y?1:x<y?-1:0)*u.dir; });
  const overdueN = S.orders.filter(isOverdue).length;
  const soonN = S.orders.filter(o=>o.status==="Rented Out"&&!isOverdue(o)&&parseD(o.ret)-d<7*DAY).length;
  const pickN = S.orders.filter(o=>o.status==="Rented Out"&&o.pickup==="BHESC Pickup").length;
  const sortTh = (k,l) => `<th class="sortable" data-act="sort" data-k="${k}">${l} ${u.sort===k?(u.dir>0?"▲":"▼"):""}</th>`;
  const allSel = list.length && list.every(o=>u.sel.has(o.no));
  return band("Orders","",`<button class="btn primary" data-act="new-order">+ Add new order</button>`) +
  tabs([["active","Open orders",counts.active],["returned","Returned & cancelled",counts.returned],["all","All orders",counts.all]], u.tab, "otab") + `
  <div class="content">
    <h2 class="sec-title">Needs attention</h2>
    <div class="opps">
      <button class="opp ${overdueN?"bad":"ok"}" data-act="ofilter" data-v="overdue"><span class="lbl">Overdue returns</span><span class="val num">${overdueN}${ico("arrow")}</span></button>
      <button class="opp warn" data-act="ofilter" data-v="soon"><span class="lbl">Due back within 7 days</span><span class="val num">${soonN}${ico("arrow")}</span></button>
      <button class="opp ${S.orders.filter(isReady).length?"warn":""}" data-act="ofilter" data-v="ready"><span class="lbl">Booked, ready to go out</span><span class="val num">${S.orders.filter(isReady).length}${ico("arrow")}</span></button>
      <button class="opp" data-act="ofilter" data-v="pickup"><span class="lbl">BHESC pickups to do</span><span class="val num">${pickN}${ico("arrow")}</span></button>
      <button class="opp ${S.orders.some(o=>cycleDue(o).length)?"warn":""}" data-act="ofilter" data-v="billdue"><span class="lbl">Cycle bills due</span><span class="val num">${S.orders.filter(o=>cycleDue(o).length).length}${ico("arrow")}</span></button>
    </div>
    ${u.filter?`<div class="infobox">Showing <b>${{overdue:"overdue returns",soon:"returns due within 7 days",ready:"booked orders ready to go out",pickup:"rented-out orders waiting on a BHESC pickup",billdue:"orders with a cycle bill due",cycle:"cycle-billed orders",booked:"booked orders only",out:"rented-out orders only"}[u.filter]}</b>. <button class="linkbtn" data-act="ofilter" data-v="">Clear</button></div>`:""}
    ${u.tab==="active"?`<div class="seg" role="group" aria-label="Show">${[["","All open",counts.active],["out","Rented out",S.orders.filter(o=>o.status==="Rented Out").length],["booked","Booked",S.orders.filter(o=>o.status==="Booked").length]].map(([k,l,n])=>`<button class="${(u.filter||"")===k||(!["out","booked"].includes(u.filter)&&k===""&&!u.filter)?"on":""}" data-act="ofilter" data-v="${k}">${l} <span class="num">${n}</span></button>`).join("")}</div>`:""}
    <div class="toolbar">
      <div class="dd"><button class="btn primary" data-act="toggle-dd" data-dd="actions">Actions ▾</button>
        <div class="dd-menu" id="dd-actions" hidden>
          <button data-act="bulk" data-v="pickup" ${u.sel.size?"":"disabled"}>Set to BHESC Pickup (${u.sel.size})</button>
          <button data-act="bulk" data-v="checkin" ${u.sel.size?"":"disabled"}>Check in selected (${u.sel.size})</button>
          <button data-act="bulk" data-v="export" ${u.sel.size?"":"disabled"}>Export selected to CSV</button>
          <button data-act="bulk" data-v="clear" ${u.sel.size?"":"disabled"}>Clear selection</button>
        </div></div>
      <div class="search">${ico("search")}<input id="oq" type="search" placeholder="Search orders" value="${esc(u.q)}" data-input="oq"></div>
      <select id="ostatus" data-change="ostatus" aria-label="Filter by status"><option value="">Filter: any status</option>${["Rented Out","Overdue","Booked","Ready to go out","Returned","Cancelled"].map(s=>`<option ${u.status===s?"selected":""}>${s}</option>`).join("")}</select>
      <div class="end"><button class="btn" data-act="export-orders">Export all to CSV</button></div>
    </div>
    <div class="tablewrap"><table class="resizable" data-cols="orders">
      <thead><tr><th style="width:36px"><input type="checkbox" data-act="selall" ${allSel?"checked":""} aria-label="Select all"></th>${sortTh("no","Order #")}${sortTh("cust","Customer")}<th>Location</th><th>Item list</th>${sortTh("price","Price")}${sortTh("out","Rent out")}${sortTh("ret","Return date")}${sortTh("status","Status")}<th>Description</th><th>Pickup status</th></tr></thead>
      <tbody>${list.length?list.map(o=>{ const c=custById(o.customerId), l=locById(o.locationId), t=orderTotals(o);
        const itemsTxt = o.lines.map(x=>{const it=itemById(x.itemId); return `#${x.itemId} - ${it?it.name:"(deleted)"}`;}).join(", ");
        return `<tr>
          <td><input type="checkbox" data-act="sel" data-no="${o.no}" ${u.sel.has(o.no)?"checked":""} aria-label="Select order ${o.no}"></td>
          <td><button class="linkbtn num" data-act="view-order" data-no="${o.no}">${o.no}</button></td>
          <td><div class="cust"><button class="linkbtn trunc" data-act="view-order" data-no="${o.no}">${esc(locLabel(l))}</button><span class="muted small">(${esc(c?c.name.split(" ").reverse().join(", "):"—")})</span></div></td>
          <td><span class="trunc">${esc(locLabel(l))}</span></td>
          <td><span class="trunc" title="${esc(itemsTxt)}">${esc(itemsTxt)}</span></td>
          <td class="num">${money(t.total)}</td>
          <td class="num">${fmtDT(o.out)}</td>
          <td class="num">${fmtDT(o.ret)}</td>
          <td>${chipFor(orderStatus(o))}${o.cycle?` <span class="chip ${cycleDue(o).length?"overdue":"neutral"}" title="Cycle billing, ${esc(cycleName(o))}">${cycleDue(o).length?"Bill due":"Cycle"}</span>`:""}</td>
          <td><span class="trunc">${esc(o.description)}</span></td>
          <td>${PICKUPS.includes(o.pickup)?chipFor(o.pickup):'<span class="muted">—</span>'}</td></tr>`;}).join(""):`<tr><td colspan="11"><div class="empty"><strong>No orders here</strong>Try another tab, clear the filters, or add a new order.</div></td></tr>`}</tbody>
    </table></div>
  </div>`;
}
function orderSearchText(o){ const c=custById(o.customerId), l=locById(o.locationId);
  return [o.no, c?.name, locLabel(l), o.description, ...o.lines.map(x=>x.itemId+" "+(itemById(x.itemId)?.name||""))].join(" ").toLowerCase(); }

/* ---------------- Items ---------------- */
function vItems(){
  const u = ui.items;
  const withSt = S.items.map(it=>({it,st:itemStatus(it)}));
  const P = {all:s=>s!=="Retired",available:s=>s==="Available",out:s=>s==="Rented Out"||s==="Overdue",co:s=>s==="Checked Out",maint:s=>s==="In Maintenance",retired:s=>s==="Retired"};
  const cnt = k => withSt.filter(x=>P[k](x.st)).length;
  let list = withSt.filter(x=>P[u.tab](x.st));
  if (u.cat){ const fam = u.exact===u.cat ? [u.cat] : groupFamily(u.cat); list = list.filter(x=>fam.includes(x.it.category)); }
  if (u.use==="rent") list = list.filter(x=>!x.it.shopOnly); else if (u.use==="shop") list = list.filter(x=>x.it.shopOnly);
  if (u.loc){ const fam = u.locExact===u.loc ? [u.loc] : locFamily(u.loc); list = list.filter(x=>fam.includes(x.it.locationId)); }
  if (u.q){ const q=u.q.toLowerCase(); list = list.filter(x=>(x.it.id+" "+x.it.name+" "+x.it.category+" "+(x.it.serial||"")+" "+homeLabel(x.it)).toLowerCase().includes(q)); }
  const cats = S.groups.map(g=>g.name).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  return band("Assets","Every asset, with live status from rentals, check-outs and the shop.",`<button class="btn" data-act="export-items">Export to CSV</button><button class="btn primary" data-act="new-item">+ Add asset</button>`) +
  tabs([["all","All assets",cnt("all")],["available","Available",cnt("available")],["out","On rent",cnt("out")],["co","Checked out",cnt("co")],["maint","In maintenance",cnt("maint")],["retired","Retired",cnt("retired")]],u.tab,"itab") + `
  <div class="content">
    ${u.cat||u.loc?`<div class="infobox">Showing ${u.cat?`group <b>${esc(groupLabel(u.cat))}</b>${u.exact===u.cat?" (not in a sub-group)":subGroups(u.cat).length?" and its sub-groups":""}`:""}${u.cat&&u.loc?" at ":""}${u.loc?`<b>${esc(locLabel(locById(u.loc)))}</b>${u.locExact===u.loc?" (not its sub-locations)":locKids(u.loc).length?" and its sub-locations":""}`:""} only. <button class="linkbtn" data-act="items-all">Show all assets</button></div>`:""}
    <div class="toolbar">
      <div class="search">${ico("search")}<input id="iq" type="search" placeholder="Search by asset # or description" value="${esc(u.q)}" data-input="iq"></div>
      <select id="icat" data-change="icat" aria-label="Group">${groupOptions(u.cat, "All groups")}</select>
      <select data-change="iuse" aria-label="Rental or shop use"><option value="">Rental and shop use</option><option value="rent" ${u.use==="rent"?"selected":""}>Rental fleet only</option><option value="shop" ${u.use==="shop"?"selected":""}>Shop use only (${S.items.filter(i=>i.shopOnly&&i.status!=="Retired").length})</option></select>
      <select id="iloc" data-change="iloc" aria-label="Location"><option value="">All locations</option>${sortedLocs().map(l=>`<option value="${l.id}" ${u.loc===l.id?"selected":""}>${esc(locOpt(l))}</option>`).join("")}</select>
    </div>
    <div class="tablewrap"><table>
      <thead><tr><th>Asset #</th><th>Description</th><th>Group</th><th style="text-align:right">Day</th><th style="text-align:right">Week</th><th style="text-align:right">Month</th><th>Status</th><th>Where / due back</th><th>Use (90 days)</th><th></th></tr></thead>
      <tbody>${list.length?list.map(({it,st})=>{ const o = activeOrderFor(it.id); const u90 = utilization(it.id);
        return `<tr><td><div class="idcell">${thumb(it,34)}<button class="linkbtn num" data-act="view-item" data-id="${esc(it.id)}">#${esc(it.id)}</button></div></td>
        <td>${esc(it.name)}${it.shopOnly?' <span class="chip shop">Shop use</span>':""}${it.notes?` <span class="chip neutral" title="${esc(it.notes)}">note</span>`:""}</td><td class="muted">${esc(groupLabel(it.category))}</td>
        <td class="num" style="text-align:right">${money(it.day)}</td><td class="num" style="text-align:right">${money(it.week)}</td><td class="num" style="text-align:right">${money(it.month)}</td>
        <td>${chipFor(st)}</td>
        <td class="small">${o?`<button class="linkbtn" data-act="view-order" data-no="${o.no}">#${o.no}</button> ${esc(locLabel(locById(o.locationId)))}<div class="muted num">Due ${fmtDate(o.ret)}</div>`:st==="In Maintenance"?`<span class="muted">${esc(openMaint(it.id).type)} · ${esc(openMaint(it.id).stage)}</span>`:st==="Checked Out"?(()=>{const c=openCustody(it.id); return `${esc(c.holder)}<div class="muted num">${c.dueAt?"Due "+fmtDate(c.dueAt):"No due date"}</div>`;})():`<span class="muted">${esc(homeLabel(it))}</span>${(()=>{ const b = nextReservation(it.id); return b ? `<div class="small"><button class="linkbtn" data-act="view-order" data-no="${b.no}">Booked #${b.no}</button> <span class="muted num">${fmtDate(b.out)} – ${fmtDate(b.ret)}</span></div>` : ""; })()}`}</td>
        <td><div class="bar-row" style="grid-template-columns:70px 36px"><div class="track"><div class="fill" style="width:${u90*100}%"></div></div><span class="v small">${Math.round(u90*100)}%</span></div></td>
        <td><div style="display:flex;gap:6px;justify-content:flex-end">${st==="Available"?`${it.shopOnly?"":`<button class="btn sm" data-act="rent-item" data-id="${esc(it.id)}">Rent</button>`}<button class="btn sm" data-act="co-item" data-id="${esc(it.id)}">Check out</button>`:""}${st==="Checked Out"?`<button class="btn sm" data-act="co-in" data-id="${openCustody(it.id).id}">Check in</button>`:""}${st==="Retired"?`<button class="btn sm" data-act="a-restore" data-id="${esc(it.id)}">Restore</button>`:""}<button class="btn sm" data-act="edit-item" data-id="${esc(it.id)}">Edit</button></div></td></tr>`;}).join(""):`<tr><td colspan="10"><div class="empty"><strong>No assets match</strong>Clear the search or pick another category.</div></td></tr>`}</tbody>
    </table></div>
  </div>`;
}

/* ---------------- Service requests (modeled on GCSS-MC: SR → tasks → labor & parts → QC → close) ---------------- */
const SR_STAGES = ["Submitted","Accepted","In work","Awaiting parts","QC inspection","Closed"];
const SR_STAGE_HINT = {"Submitted":"Reported, not in the shop yet","Accepted":"Received in the shop (DRIS)","In work":"Mechanics working tasks","Awaiting parts":"Held for parts on order","QC inspection":"Repairs done, waiting on final inspection","Closed":"Passed QC, back in service"};
const REQ_TYPES = ["Corrective maintenance","Preventive maintenance","Inspection","Modification","Warranty"];
const OP_STATUS = [["Operational","Still usable. Stays rentable until it's accepted into the shop"],["Degraded","Usable with limits until it's repaired"],["Deadlined","Down. Can't be used or rented"]];
const PROBLEM_CODES = [["ENG","Engine"],["HYD","Hydraulics"],["ELE","Electrical"],["DRV","Drivetrain / transmission"],["BRK","Brakes"],["TRK","Tires / tracks / undercarriage"],["STE","Steering"],["BOD","Body / frame / structure"],["ATT","Attachment / implement"],["LEK","Fluid leak"],["SAF","Safety item (lights, alarms, ROPS)"],["PMS","Scheduled service"],["OTH","Other"]];
const TASK_STATUS = ["Open","In work","Awaiting parts","Complete","Cancelled"];
const PART_STATUS = ["Needed","Ordered","Received","Installed","Cancelled"];
const SR_CHIPS = {"Pending":"maint","Approved":"available","Declined":"overdue","Submitted":"neutral","Accepted":"checked","In work":"rented","Awaiting parts":"maint","QC inspection":"reserved","Closed":"available","Operational":"available","Degraded":"maint","Deadlined":"overdue","Needed":"overdue","Ordered":"maint","Received":"rented","Installed":"available","Complete":"available"};
const srNo = m => m && m.no ? "SR-" + m.no : "SR";
const srById = id => S.maintenance.find(m=>m.id===id);
const codeName = c => (PROBLEM_CODES.find(x=>x[0]===c)||[c,c])[1];
const laborRate = () => +(S.settings.charges.find(c=>/mechanic/i.test(c.name))?.rate || 0);
function srLabor(m){ return (m.tasks||[]).reduce((s,t)=>s+(t.labor||[]).reduce((a,l)=>a+(+l.hours||0),0),0); }
function srParts(m){ return (m.tasks||[]).flatMap(t=>t.parts||[]); }
function srOpenParts(m){ return srParts(m).filter(p=>p.status==="Needed"||p.status==="Ordered"); }
function partsCost(m){ return (m.parts||[]).reduce((s,p)=>s+(+p.qty||0)*(+p.cost||0),0) + srParts(m).filter(p=>p.status!=="Cancelled").reduce((s,p)=>s+(+p.qty||0)*(+p.cost||0),0); }
function srCost(m){ return partsCost(m) + (+m.cost||0) + (m.tasks ? srLabor(m) : (+m.hours||0)) * laborRate(); }
function srDays(m){ const a = parseD(((m.acceptedAt||m.date)+"").slice(0,10)+"T00:00"), b = m.doneAt ? parseD(m.doneAt.slice(0,10)+"T00:00") : now(); return Math.max(0, Math.floor((b-a)/DAY)); }
function srLog(m, text){ (m.log ||= []).push({at:toLocalInput(now()), by:myName(), text}); }
function srTarget(){ return +S.settings.shopTargetDays || 30; }
/* keep derived fields, inventory stock and the shop stage in step with the tasks */
function srSync(m){
  if (!m.tasks) return;
  m.hours = +srLabor(m).toFixed(2);
  const owners = [...new Set(m.tasks.filter(t=>t.status!=="Cancelled").map(t=>t.owner).filter(Boolean))]; m.tech = owners.join(", ");
  const at = toLocalInput(now());
  m.tasks.forEach(t=>(t.parts||[]).forEach(p=>{ if (!p.invId) return; const inv = S.inventory.find(x=>x.id===p.invId); if (!inv) return;
    const want = (p.status==="Received"||p.status==="Installed") ? (+p.qty||0) : 0, d = +(want - (+p.taken||0)).toFixed(2); if (!d) return;
    inv.qty = +(inv.qty - d).toFixed(2); (inv.log ||= []).push({by:myName(), at, delta:-d, reason: d>0 ? `Issued to ${srNo(m)} (#${m.itemId})` : `Returned to stock from ${srNo(m)}`}); p.taken = want; }));
  if (m.status==="Open"){
    const op = srOpenParts(m).length;
    if (op && ["Accepted","In work"].includes(m.stage)){ m.stage = "Awaiting parts"; srLog(m, "Status → Awaiting parts (parts on order)"); }
    else if (!op && m.stage==="Awaiting parts"){ m.stage = "In work"; srLog(m, "Status → In work (all parts in)"); }
  }
  if (m.opStatus==="Deadlined"){ if (!m.deadlineAt && m.status==="Open") m.deadlineAt = at; } else if (m.status==="Open") m.deadlineAt = null;
}
function newSR(f){
  const m = {id:uid("m"), no:S.settings.nextSR++, itemId:f.itemId, date:f.date||toLocalInput(now()).slice(0,10), type:f.type||"Corrective maintenance", severity:f.severity||"", opStatus:f.opStatus||"Operational",
    codes:f.codes||[], desc:f.desc||"", reportedBy:f.reportedBy||myName(), stage:f.stage||"Submitted", status:"Open", acceptedAt:null, deadlineAt:null, doneAt:null, qc:null, cost:0, hours:0, tech:"", parts:[], files:f.files||[], log:[],
    tasks:[{id:uid("t"), subject:f.task || (f.desc||"Inspect and repair").split("\n")[0].slice(0,90), code:(f.codes||[])[0]||"", status:"Open", owner:f.owner||"", planned:"", labor:[], parts:[]}]};
  if (m.stage!=="Submitted") m.acceptedAt = m.date;
  srLog(m, `Service request submitted${f.why?` (${f.why})`:""}`);
  srSync(m); S.maintenance.push(m); return m;
}
function srReadyToClose(m){
  const why = [];
  const openT = (m.tasks||[]).filter(t=>t.status!=="Complete"&&t.status!=="Cancelled"); if (openT.length) why.push(`${openT.length} task${openT.length>1?"s":""} not complete`);
  const op = srOpenParts(m); if (op.length) why.push(`${op.length} part${op.length>1?"s":""} still needed or on order`);
  const rec = srParts(m).filter(p=>p.status==="Received"); if (rec.length) why.push(`${rec.length} received part${rec.length>1?"s":""} not marked installed`);
  return why;
}
function closeSR(m, by, note){
  const at = toLocalInput(now());
  m.qc = {by, at, note:note||"Repairs inspected, QC complete", passed:true};
  m.stage = "Closed"; m.status = "Completed"; m.doneAt = at.slice(0,10); m.opStatus = "Operational";
  srLog(m, `Passed QC by ${by}${note?`: ${note}`:""}. Closed, back in service`); srSync(m);
}

/* ---- New / edit SR header ---- */
function maintForm(id, itemId){
  const m = id ? srById(id) : {itemId:itemId||S.items.find(i=>i.status!=="Retired")?.id, date:toLocalInput(now()).slice(0,10), type:"Corrective maintenance", severity:"Low", opStatus:"Operational", codes:[], desc:"", reportedBy:myName(), stage:"Submitted"};
  mdraft = {id:id||null, files:(m.files||[]).slice(), codes:(m.codes||[]).slice()};
  const users = S.users.filter(u=>u.active!==false).map(u=>u.name);
  openModal(`<header><h2>${id?`Edit ${srNo(m)}`:"New service request"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="maintform" class="form" data-id="${id||""}">
    <div class="field full"><label for="m-item">Asset</label><select id="m-item" name="itemId" ${id?"disabled":""}>${S.items.filter(i=>i.status!=="Retired"||i.id===m.itemId).map(i=>`<option value="${esc(i.id)}" ${i.id===m.itemId?"selected":""}>#${esc(i.id)} ${esc(i.name)}</option>`).join("")}</select></div>
    <div class="field"><label for="m-type">Request type</label><select id="m-type" name="type">${REQ_TYPES.concat(REQ_TYPES.includes(m.type)?[]:[m.type]).map(t=>`<option ${t===m.type?"selected":""}>${esc(t)}</option>`).join("")}</select></div>
    <div class="field"><label for="m-date">Date reported</label><input id="m-date" type="date" name="date" value="${esc(m.date)}"></div>
    <div class="field"><label for="m-sev">Severity</label><select id="m-sev" name="severity">${m.severity?"":`<option value="" selected>Not set</option>`}${SEVERITY.map(([v,d])=>`<option value="${v}" ${v===m.severity?"selected":""}>${v} – ${d.split(" – ")[0]}</option>`).join("")}</select><span class="hint" id="m-sev-hint">${esc((SEVERITY.find(x=>x[0]===m.severity)||[,""])[1].split(" – ")[1]||"")}</span></div>
    <div class="field"><label for="m-op">Operational status</label><select id="m-op" name="opStatus">${OP_STATUS.map(([v])=>`<option ${v===m.opStatus?"selected":""}>${v}</option>`).join("")}</select><span class="hint" id="m-op-hint">${esc((OP_STATUS.find(x=>x[0]===m.opStatus)||OP_STATUS[0])[1])}</span></div>
    <div class="field full"><label>Problem codes</label><div class="code-pick" id="m-codes">${PROBLEM_CODES.map(([c,n])=>`<label class="code-opt"><input type="checkbox" data-code="${c}" ${mdraft.codes.includes(c)?"checked":""}><span><b>${c}</b> ${esc(n)}</span></label>`).join("")}</div></div>
    <div class="field full"><label for="m-desc">Problem summary</label><textarea id="m-desc" name="desc" placeholder="What's wrong, what the operator saw, when it happened">${esc(m.desc)}</textarea></div>
    <div class="field"><label for="m-rep">Reported by</label><input id="m-rep" name="reportedBy" value="${esc(m.reportedBy||"")}" list="m-people"></div>
    ${id?"":`<div class="field"><label for="m-owner">Assign to mechanic</label><input id="m-owner" name="owner" list="m-people" placeholder="Optional"></div>
    <div class="field full"><label class="code-opt" style="border:0;padding:0"><input type="checkbox" id="m-accept" name="accept" ${itemId?"checked":""}><span>Asset is already in the shop: <b>accept it now</b></span></label></div>`}
    <datalist id="m-people">${[...new Set(users.concat(S.maintenance.flatMap(x=>(x.tasks||[]).map(t=>t.owner)).filter(Boolean)))].map(n=>`<option value="${esc(n)}">`).join("")}</datalist>
  </form>
  <div class="subhead"><span>Documentation</span><button class="btn sm" data-act="m-addfile">${ico("clip")} Attach files</button></div>
  <div id="m-files" class="dropzone" data-drop="mfiles">${docsEditor()}</div>
  </div>
  <footer>${id&&can("admin")?`<button class="btn danger" data-act="del-maint" data-id="${id}" style="margin-right:auto">Delete</button>`:""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="save-maint">${id?"Save":"Submit request"}</button></footer>`,"narrow");
}
function saveMaint(){
  const f = $("#maintform"); const fd = Object.fromEntries(new FormData(f));
  const codes = [...document.querySelectorAll("#m-codes input:checked")].map(x=>x.dataset.code);
  if (!fd.desc.trim()){ toast("Add a problem summary"); $("#m-desc").focus(); return; }
  if (f.dataset.id){
    const m = srById(f.dataset.id), ch = [];
    if (m.opStatus!==fd.opStatus) ch.push(`Operational status ${m.opStatus} → ${fd.opStatus}`);
    if ((m.severity||"")!==fd.severity) ch.push(`Severity ${m.severity||"not set"} → ${fd.severity||"not set"}`);
    if (m.type!==fd.type) ch.push(`Type → ${fd.type}`);
    Object.assign(m, {type:fd.type, date:fd.date, severity:fd.severity, opStatus:fd.opStatus, desc:fd.desc, reportedBy:fd.reportedBy, codes, files:mdraft.files});
    srLog(m, ch.length ? ch.join("; ") : "Details edited"); srSync(m); mdraft = null; save(); viewSR(m.id); rerender(); toast(`${srNo(m)} saved`); return;
  }
  const m = newSR({itemId:fd.itemId, date:fd.date, type:fd.type, severity:fd.severity, opStatus:fd.opStatus, codes, desc:fd.desc, reportedBy:fd.reportedBy, owner:fd.owner, files:mdraft.files, stage: fd.accept ? "Accepted" : "Submitted"});
  if (fd.accept) srLog(m, "Accepted into the shop");
  mdraft = null; save(); rerender(); viewSR(m.id); toast(`${srNo(m)} submitted`);
}

/* ---- SR detail ---- */
let srUI = {id:null, addTask:false, labor:null, part:null, qc:false};
function viewSR(id){ srUI = {id, addTask:false, labor:null, part:null, qc:false}; drawSR(true); }
function drawSR(first){
  const m = srById(srUI.id); if (!m) { closeModal(); return; }
  const it = itemById(m.itemId), edit = can("shop") && m.status==="Open", days = srDays(m), late = m.status==="Open" && days > srTarget();
  const idx = SR_STAGES.indexOf(m.stage);
  const steps = SR_STAGES.map((s,i)=>`<li class="${i<idx?"done":i===idx?"on":""}" title="${esc(SR_STAGE_HINT[s])}"><span>${i<idx?"✓":i+1}</span>${s}</li>`).join("");
  const users = [...new Set(S.users.filter(u=>u.active!==false).map(u=>u.name).concat(S.maintenance.flatMap(x=>(x.tasks||[]).map(t=>t.owner)).filter(Boolean)))];
  const blockers = srReadyToClose(m);
  let next = "";
  if (can("shop")){
    if (m.stage==="Submitted") next = `<button class="btn primary" data-act="sr-accept">Accept into shop</button>`;
    else if (m.stage==="Accepted") next = `<button class="btn primary" data-act="sr-stage" data-v="In work">Start work</button>`;
    else if (m.stage==="In work" || m.stage==="Awaiting parts") next = `<button class="btn primary" data-act="sr-stage" data-v="QC inspection" ${blockers.length?`disabled title="${esc(blockers.join(", "))}"`:""}>Send to QC</button>`;
    else if (m.stage==="QC inspection") next = `<button class="btn primary" data-act="sr-qc-open">Final inspection</button>`;
    else if (m.stage==="Closed") next = `<button class="btn" data-act="sr-reopen">Reopen</button>`;
  }
  const taskHtml = (m.tasks||[]).map((t,ti)=>{ const hrs = (t.labor||[]).reduce((a,l)=>a+(+l.hours||0),0);
    const lab = (t.labor||[]).map((l,li)=>`<tr><td class="num">${fmtDate(l.date+"T00:00")}</td><td>${esc(l.who)}</td><td class="num" style="text-align:right">${l.hours}</td><td class="muted">${esc(l.note||"")}</td><td>${edit?`<button class="del" data-act="sr-labor-del" data-t="${ti}" data-i="${li}" aria-label="Remove labor">×</button>`:""}</td></tr>`).join("");
    const prt = (t.parts||[]).map((p,pi)=>`<tr><td><b>${esc(p.name)}</b>${p.partNo?`<div class="small muted num">${esc(p.partNo)}</div>`:""}</td><td class="small">${p.invId?"From inventory":esc(p.vendor||"Vendor")}${p.po?`<div class="muted">PO ${esc(p.po)}</div>`:""}</td><td class="num" style="text-align:right">${p.qty}</td><td class="num" style="text-align:right">${money((+p.qty||0)*(+p.cost||0))}</td>
      <td>${edit?`<select class="sm-select" data-change="sr-part-status" data-t="${ti}" data-i="${pi}" aria-label="Part status">${PART_STATUS.map(s=>`<option ${s===p.status?"selected":""}>${s}</option>`).join("")}</select>`:chipFor(p.status)}</td><td>${edit?`<button class="del" data-act="sr-part-del" data-t="${ti}" data-i="${pi}" aria-label="Remove part">×</button>`:""}</td></tr>`).join("");
    const lf = srUI.labor===ti, pf = srUI.part===ti;
    return `<div class="sr-task ${t.status==="Complete"?"is-done":t.status==="Cancelled"?"is-cancel":""}">
      <div class="sr-task-h"><span class="sr-task-n">Task ${ti+1}</span>
        ${edit?`<input class="sr-subj" value="${esc(t.subject)}" data-change="sr-task" data-t="${ti}" data-f="subject" aria-label="Task">`:`<b class="sr-subj">${esc(t.subject)}</b>`}
        ${edit?`<select class="sm-select" data-change="sr-task" data-t="${ti}" data-f="status" aria-label="Task status">${TASK_STATUS.map(s=>`<option ${s===t.status?"selected":""}>${s}</option>`).join("")}</select>`:chipFor(t.status)}</div>
      <div class="sr-task-meta">
        <label>Code ${edit?`<select class="sm-select" data-change="sr-task" data-t="${ti}" data-f="code"><option value="">—</option>${PROBLEM_CODES.map(([c,n])=>`<option value="${c}" ${c===t.code?"selected":""}>${c} · ${esc(n)}</option>`).join("")}</select>`:`<b>${t.code?esc(t.code+" · "+codeName(t.code)):"—"}</b>`}</label>
        <label>Mechanic ${edit?`<input class="sm-input" list="sr-people" value="${esc(t.owner||"")}" data-change="sr-task" data-t="${ti}" data-f="owner" placeholder="Unassigned">`:`<b>${esc(t.owner||"Unassigned")}</b>`}</label>
        <label>Planned hrs ${edit?`<input class="sm-input num" type="number" step="0.25" min="0" style="width:70px" value="${t.planned||""}" data-change="sr-task" data-t="${ti}" data-f="planned">`:`<b>${t.planned||"—"}</b>`}</label>
        <span class="small">Actual <b class="num ${t.planned&&hrs>+t.planned?"bad":""}">${+hrs.toFixed(2)} hrs</b></span>
        ${edit&&m.tasks.length>1&&!hrs&&!(t.parts||[]).length?`<button class="linkbtn small" data-act="sr-task-del" data-t="${ti}" style="margin-left:auto;color:var(--bad)">Remove task</button>`:""}
      </div>
      <div class="sr-sub"><div class="sr-sub-h"><span>Labor</span>${edit&&!lf?`<button class="btn sm" data-act="sr-labor-new" data-t="${ti}">+ Log labor</button>`:""}</div>
        ${lab||lf?`<table class="sr-mini"><tbody>${lab}${lf?`<tr class="sr-form-row"><td><input type="date" id="lb-date" value="${toLocalInput(now()).slice(0,10)}" aria-label="Date"></td><td><input id="lb-who" list="sr-people" value="${esc(t.owner||myName())}" aria-label="Mechanic"></td><td><input id="lb-hrs" type="number" step="0.25" min="0" placeholder="Hrs" aria-label="Hours" style="width:70px"></td><td><input id="lb-note" placeholder="What was done" aria-label="Note"></td><td><div style="display:flex;gap:4px"><button class="btn sm primary" data-act="sr-labor-add" data-t="${ti}">Log</button><button class="btn sm" data-act="sr-form-cancel">×</button></div></td></tr>`:""}</tbody></table>`:`<div class="small muted">No labor logged.</div>`}</div>
      <div class="sr-sub"><div class="sr-sub-h"><span>Parts</span>${edit&&!pf?`<button class="btn sm" data-act="sr-part-new" data-t="${ti}">+ Add part</button>`:""}</div>
        ${prt||pf?`<table class="sr-mini"><tbody>${prt}${pf?`<tr class="sr-form-row"><td colspan="6"><div class="sr-partform">
          <select id="pt-src" data-change="sr-part-src" aria-label="Source"><option value="">Order from a vendor</option>${S.inventory.filter(x=>!x.retired).map(x=>`<option value="${x.id}">From inventory: ${esc(x.name)} (${x.qty} ${esc(x.unit||"")})</option>`).join("")}</select>
          <input id="pt-name" placeholder="Part name" aria-label="Part name"><input id="pt-no" placeholder="Part #" aria-label="Part number"><input id="pt-vendor" placeholder="Vendor" aria-label="Vendor"><input id="pt-po" placeholder="PO #" aria-label="PO number">
          <input id="pt-qty" type="number" step="0.1" min="0" value="1" aria-label="Quantity"><input id="pt-cost" type="number" step="0.01" min="0" placeholder="Unit cost" aria-label="Unit cost">
          <select id="pt-status" aria-label="Status">${PART_STATUS.filter(s=>s!=="Cancelled").map(s=>`<option ${s==="Ordered"?"selected":""}>${s}</option>`).join("")}</select>
          <button class="btn sm primary" data-act="sr-part-add" data-t="${ti}">Add</button><button class="btn sm" data-act="sr-form-cancel">×</button></div></td></tr>`:""}</tbody></table>`:`<div class="small muted">No parts.</div>`}</div>
    </div>`; }).join("");
  const qcHtml = m.qc ? `<div class="sr-qc ok"><span class="sr-qc-ic">✓</span><div><b>Passed final inspection</b> · ${esc(m.qc.by||"")} · ${fmtDate(m.qc.at)}${m.qc.note?`<div class="small">${esc(m.qc.note)}</div>`:""}</div></div>` : srUI.qc ? `<div class="sr-qc">
      <div class="form" style="grid-template-columns:1fr 1fr"><div class="field"><label for="qc-by">Inspected by</label><input id="qc-by" list="sr-people" value="${esc(myName())}"></div><div class="field"><label for="qc-meter">Hour meter at close (optional)</label><input id="qc-meter" type="number" step="0.1" value="${it?.hours??""}"></div>
      <div class="field full"><label for="qc-note">Inspection notes</label><input id="qc-note" value="Repairs inspected, operational check complete"></div></div>
      <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px"><button class="btn" data-act="sr-qc-fail">Failed, back to work</button><button class="btn primary" data-act="sr-qc-pass">Passed, close SR</button></div></div>` : "";
  const html = `<header><h2>${srNo(m)} · #${esc(m.itemId)} ${esc(it?.name||"")}</h2>${chipFor(m.stage)}${chipFor(m.opStatus||"Operational")}<button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody" id="sr-body">
    <ol class="sr-steps">${steps}</ol>
    <div class="sr-grid">
      <div><span>Type</span><b>${esc(m.type)}</b></div><div><span>Severity</span><b>${sevChip(m.severity)}</b></div>
      <div><span>Operational status</span>${edit?`<select class="sm-select" data-change="sr-op" aria-label="Operational status">${OP_STATUS.map(([v])=>`<option ${v===m.opStatus?"selected":""}>${v}</option>`).join("")}</select>`:`<b>${esc(m.opStatus||"Operational")}</b>`}</div>
      <div><span>Reported</span><b>${fmtDate(m.date+"T00:00")}${m.reportedBy?` · ${esc(m.reportedBy)}`:""}</b></div>
      <div><span>In shop since (DRIS)</span><b>${m.acceptedAt?fmtDate(m.acceptedAt.slice(0,10)+"T00:00"):"Not accepted yet"}</b></div>
      <div><span>Days in shop</span><b class="num ${late?"bad":""}">${m.acceptedAt?days:"—"}${late?` · over ${srTarget()}-day target`:""}</b></div>
      ${m.deadlineAt?`<div><span>Deadlined since</span><b class="bad">${fmtDate(m.deadlineAt)}</b></div>`:""}
      <div><span>Labor</span><b class="num">${+srLabor(m).toFixed(2)} hrs · ${money(srLabor(m)*laborRate())}</b></div>
      <div><span>Parts</span><b class="num">${money(partsCost(m))}${srOpenParts(m).length?` · <span style="color:var(--warn)">${srOpenParts(m).length} open</span>`:""}</b></div>
      <div><span>Total cost</span><b class="num">${money(srCost(m))}</b></div>
    </div>
    ${(m.codes||[]).length?`<div class="sr-codes">${m.codes.map(c=>`<span class="chip neutral">${esc(c)} · ${esc(codeName(c))}</span>`).join("")}</div>`:""}
    <div class="sr-summary">${esc(m.desc||"No problem summary.")}</div>
    ${m.stage==="QC inspection"||m.qc?`<div class="subhead"><span>Final inspection (QC)</span></div>${qcHtml||`<div class="small muted">Ready for final inspection.</div>`}`:""}
    <div class="subhead"><span>Tasks <span class="muted">${(m.tasks||[]).filter(t=>t.status==="Complete").length}/${(m.tasks||[]).filter(t=>t.status!=="Cancelled").length} complete</span></span>${edit&&!srUI.addTask?`<button class="btn sm" data-act="sr-task-new">+ Add task</button>`:""}</div>
    ${srUI.addTask?`<div class="sr-task sr-newtask"><div class="sr-partform"><input id="tk-subj" placeholder="Task, e.g. Replace hydraulic hose" aria-label="Task" style="flex:2"><select id="tk-code" aria-label="Problem code"><option value="">Code</option>${PROBLEM_CODES.map(([c,n])=>`<option value="${c}">${c} · ${esc(n)}</option>`).join("")}</select><input id="tk-owner" list="sr-people" placeholder="Mechanic" aria-label="Mechanic"><input id="tk-plan" type="number" step="0.25" min="0" placeholder="Planned hrs" aria-label="Planned hours" style="width:100px"><button class="btn sm primary" data-act="sr-task-add">Add task</button><button class="btn sm" data-act="sr-form-cancel">×</button></div></div>`:""}
    <div class="sr-tasks">${taskHtml||'<div class="muted small">No tasks yet.</div>'}</div>
    <div class="subhead"><span>Documentation</span>${can("shop")?`<button class="btn sm" data-act="sr-addfile">${ico("clip")} Attach files</button>`:""}</div>
    <div class="files">${(m.files||[]).map((f,i)=>docRow(f,i,false,m.id)).join("")||'<div class="small muted">No documents.</div>'}</div>
    <div class="subhead"><span>History</span></div>
    <div class="sr-log">${(m.log||[]).slice().reverse().map(l=>`<div><span class="num muted">${fmtDate(l.at)}</span> <b>${esc(l.by||"")}</b> ${esc(l.text)}</div>`).join("")||'<div class="small muted">No history.</div>'}</div>
    <datalist id="sr-people">${users.map(n=>`<option value="${esc(n)}">`).join("")}</datalist>
  </div>
  <footer>${can("shop")?`<button class="btn" data-act="edit-maint" data-id="${m.id}" style="margin-right:auto">Edit details</button>`:`<span style="margin-right:auto"></span>`}
    <button class="btn" data-act="sr-print">${ico("download")} Print / save</button><button class="btn" data-act="close">Close</button>${next}</footer>`;
  if (first || !$("#sr-body")) openModal(html, "sr");
  else { const y = $("#sr-body").scrollTop; $("#modal-root .modal").innerHTML = html; applyPerms($("#modal-root")); $("#sr-body").scrollTop = y; }
}
function srChange(fn, msg){ const m = srById(srUI.id); fn(m); srSync(m); save(); drawSR(); rerender(); if (msg) toast(msg); }
function srAction(act, ds){
  const m = srById(srUI.id); if (!m) return false;
  const v = id => ($("#"+id)?.value ?? "").trim();
  switch(act){
    case "sr-accept": srChange(m=>{ m.stage = "Accepted"; m.acceptedAt = toLocalInput(now()).slice(0,10); srLog(m, "Accepted into the shop (date received in shop set)"); }, `${srNo(m)} accepted. #${m.itemId} is out of service`); return true;
    case "sr-stage": srChange(m=>{ m.stage = ds.v; if (!m.acceptedAt) m.acceptedAt = toLocalInput(now()).slice(0,10); if (ds.v==="In work") m.tasks.forEach(t=>{ if (t.status==="Open") t.status = "In work"; }); srLog(m, `Status → ${ds.v}`); }); return true;
    case "sr-qc-open": srUI.qc = true; drawSR(); return true;
    case "sr-qc-pass": { const why = srReadyToClose(m); if (why.length){ toast("Can't close yet: " + why.join(", ")); return true; }
      const by = v("qc-by") || myName(), meter = v("qc-meter"); const it = itemById(m.itemId);
      srChange(m=>{ if (meter !== "" && it && +meter !== +it.hours){ it.hours = +meter; } closeSR(m, by, v("qc-note")); }, `${srNo(m)} closed. #${m.itemId} is back in service`); srUI.qc = false; drawSR(); return true; }
    case "sr-qc-fail": { const note = v("qc-note"); srChange(m=>{ m.stage = "In work"; srLog(m, `Failed final inspection${note?`: ${note}`:""}. Back to work`); m.tasks.push({id:uid("t"), subject:"Correct QC findings" + (note?`: ${note}`:""), code:"", status:"Open", owner:"", planned:"", labor:[], parts:[]}); }, "Sent back to work with a new task"); srUI.qc = false; drawSR(); return true; }
    case "sr-reopen": srChange(m=>{ m.status = "Open"; m.stage = "In work"; m.doneAt = null; m.qc = null; srLog(m, "Reopened"); }, `${srNo(m)} reopened`); return true;
    case "sr-task-new": srUI.addTask = true; srUI.labor = srUI.part = null; drawSR(); $("#tk-subj")?.focus(); return true;
    case "sr-task-add": { const subj = v("tk-subj"); if (!subj){ toast("Name the task"); return true; }
      srChange(m=>{ m.tasks.push({id:uid("t"), subject:subj, code:v("tk-code"), status:"Open", owner:v("tk-owner"), planned:v("tk-plan"), labor:[], parts:[]}); srLog(m, `Task added: ${subj}`); }); srUI.addTask = false; drawSR(); return true; }
    case "sr-task-del": srChange(m=>{ const [t] = m.tasks.splice(+ds.t,1); srLog(m, `Task removed: ${t.subject}`); }); return true;
    case "sr-form-cancel": srUI.addTask = false; srUI.labor = srUI.part = null; drawSR(); return true;
    case "sr-labor-new": srUI.labor = +ds.t; srUI.part = null; srUI.addTask = false; drawSR(); $("#lb-hrs")?.focus(); return true;
    case "sr-labor-add": { const h = +v("lb-hrs"); if (!(h>0)){ toast("Enter the hours worked"); return true; }
      srChange(m=>{ const t = m.tasks[+ds.t]; t.labor.push({id:uid("l"), who:v("lb-who"), date:v("lb-date")||toLocalInput(now()).slice(0,10), hours:h, note:v("lb-note")}); if (t.status==="Open") t.status = "In work"; if (m.stage==="Accepted") { m.stage = "In work"; srLog(m, "Status → In work"); } srLog(m, `${h} hrs logged on "${t.subject}" by ${v("lb-who")||"—"}`); }); srUI.labor = null; drawSR(); return true; }
    case "sr-labor-del": srChange(m=>{ const t = m.tasks[+ds.t]; const [l] = t.labor.splice(+ds.i,1); srLog(m, `Removed ${l.hours} hrs from "${t.subject}"`); }); return true;
    case "sr-part-new": srUI.part = +ds.t; srUI.labor = null; srUI.addTask = false; drawSR(); $("#pt-name")?.focus(); return true;
    case "sr-part-add": { const src = v("pt-src"), inv = S.inventory.find(x=>x.id===src), name = inv ? inv.name : v("pt-name"), qty = +v("pt-qty");
      if (!name){ toast("Name the part"); return true; } if (!(qty>0)){ toast("Enter a quantity"); return true; }
      srChange(m=>{ const t = m.tasks[+ds.t]; t.parts.push({id:uid("p"), invId:inv?inv.id:"", name, partNo:v("pt-no")||(inv?.sku||""), vendor:inv?"":v("pt-vendor"), po:v("pt-po"), qty, cost:+v("pt-cost")||+(inv?.unitCost ?? 0)||0, status:v("pt-status")||"Ordered", taken:0}); srLog(m, `Part ${v("pt-status")||"Ordered"}: ${qty} × ${name}${v("pt-po")?` (PO ${v("pt-po")})`:""}`); }); srUI.part = null; drawSR(); return true; }
    case "sr-part-del": srChange(m=>{ const t = m.tasks[+ds.t], p = t.parts[+ds.i]; p.status = "Cancelled"; srSync(m); t.parts.splice(+ds.i,1); srLog(m, `Part removed: ${p.name}`); }); return true;
    case "sr-addfile": pickFiles("", true, async files=>{ mdraft = {files:(m.files||[]).slice(), live:m.id}; await addDocs(files); }); return true;
    case "sr-print": { const doc = srDoc(m); if (IN_VIEWER) saveFile(`${srNo(m)}-${slug(itemById(m.itemId)?.name||"asset")}.html`, srDoc(m, true), "text/html;charset=utf-8"); else printHtmlDoc(doc); return true; }
  }
  return false;
}
function srFieldChange(t){
  const ds = t.dataset, m = srById(srUI.id); if (!m) return false;
  if (ds.change==="sr-op"){ srChange(m=>{ srLog(m, `Operational status ${m.opStatus} → ${t.value}`); m.opStatus = t.value; }); return true; }
  if (ds.change==="sr-task"){ const tk = m.tasks[+ds.t], f = ds.f, val = t.value; if ((tk[f]??"")===val) return true;
    srChange(m=>{ tk[f] = f==="planned" ? (val===""?"":+val) : val; if (f==="status") srLog(m, `Task "${tk.subject}" → ${val}`); else if (f==="owner") srLog(m, `Task "${tk.subject}" assigned to ${val||"nobody"}`); }); return true; }
  if (ds.change==="sr-part-status"){ const p = m.tasks[+ds.t].parts[+ds.i]; srChange(m=>{ srLog(m, `Part ${p.name}: ${p.status} → ${t.value}`); p.status = t.value; }); return true; }
  if (ds.change==="sr-part-src"){ const inv = S.inventory.find(x=>x.id===t.value); ["pt-name","pt-vendor","pt-po"].forEach(id=>{ const el = $("#"+id); if (el){ el.disabled = !!inv; if (inv && id==="pt-name") el.value = inv.name; } });
    if (inv){ if ($("#pt-no")) $("#pt-no").value = inv.sku||""; if ($("#pt-cost")) $("#pt-cost").value = inv.unitCost ?? ""; if ($("#pt-status")) $("#pt-status").value = "Installed"; } else if ($("#pt-status")) $("#pt-status").value = "Ordered"; return true; }
  return false;
}
function printHtmlDoc(doc){ const f = document.createElement("iframe"); f.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0"; document.body.appendChild(f); f.srcdoc = doc;
  f.onload = () => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch(e){ toast("Printing is blocked here."); } setTimeout(()=>f.remove(), 60000); }; }
function srDoc(m, autoPrint){
  const it = itemById(m.itemId), e = esc;
  const tasks = (m.tasks||[]).map((t,i)=>`<tr><td>${i+1}</td><td><b>${e(t.subject)}</b>${t.code?`<div class="sub">${e(t.code)} · ${e(codeName(t.code))}</div>`:""}</td><td>${e(t.owner||"")}</td><td>${e(t.status)}</td><td class="r">${t.planned||""}</td><td class="r">${+(t.labor||[]).reduce((a,l)=>a+(+l.hours||0),0).toFixed(2)}</td></tr>`).join("");
  const parts = srParts(m).filter(p=>p.status!=="Cancelled").map(p=>`<tr><td>${e(p.name)}</td><td>${e(p.partNo||"")}</td><td>${p.invId?"Inventory":e(p.vendor||"")}${p.po?` · PO ${e(p.po)}`:""}</td><td>${e(p.status)}</td><td class="r">${p.qty}</td><td class="r">${money((+p.qty||0)*(+p.cost||0))}</td></tr>`).join("");
  const labor = (m.tasks||[]).flatMap(t=>(t.labor||[]).map(l=>`<tr><td>${fmtDate(l.date+"T00:00")}</td><td>${e(l.who)}</td><td>${e(t.subject)}</td><td>${e(l.note||"")}</td><td class="r">${l.hours}</td></tr>`)).join("");
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${srNo(m)} · #${e(m.itemId)} ${e(it?.name||"")}</title><style>
  @page{size:letter;margin:.5in}*{box-sizing:border-box}body{font:11px/1.45 Arial,Helvetica,sans-serif;color:#16170f;margin:0}
  .doc{max-width:7.5in;margin:0 auto}h1{font-size:20px;margin:0}h2{font-size:12px;text-transform:uppercase;letter-spacing:.06em;border-bottom:2px solid #16170f;padding-bottom:3px;margin:18px 0 6px}
  .top{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #4b5320;padding-bottom:10px}.top img{max-height:44px}
  .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:6px 14px;margin-top:10px}.grid div span{display:block;font-size:9px;text-transform:uppercase;color:#62665a;letter-spacing:.05em}
  table{width:100%;border-collapse:collapse}th{text-align:left;font-size:9px;text-transform:uppercase;color:#62665a;border-bottom:1px solid #16170f;padding:4px}td{padding:5px 4px;border-bottom:1px solid #dfe1d4;vertical-align:top}.r{text-align:right}.sub{color:#62665a;font-size:10px}
  .box{border:1px solid #c9ccbd;padding:8px;min-height:40px;white-space:pre-wrap}.sig{display:grid;grid-template-columns:2fr 1fr;gap:24px;margin-top:30px}.sig span{display:block;border-bottom:1px solid #16170f;height:24px}
  @media screen{body{background:#e9eadf}.doc{background:#fff;padding:.5in;margin:16px auto;box-shadow:0 2px 14px rgba(0,0,0,.15)}}</style></head><body><div class="doc">
  <div class="top"><div>${S.settings.logo?`<img src="${S.settings.logo}" alt=""><br>`:""}<b>${e(S.settings.company)}</b><div class="sub">Maintenance service request</div></div><div style="text-align:right"><h1>${srNo(m)}</h1><div>${e(m.stage)} · ${e(m.opStatus||"")}</div><div class="sub">Printed ${fmtDate(toLocalInput(now()))}</div></div></div>
  <div class="grid"><div><span>Asset</span>#${e(m.itemId)} ${e(it?.name||"")}</div><div><span>Serial / VIN</span>${e(it?.serial||it?.vin||"—")}</div><div><span>Type</span>${e(m.type)}</div><div><span>Severity</span>${e(m.severity||"—")}</div>
  <div><span>Reported</span>${fmtDate(m.date+"T00:00")} ${e(m.reportedBy||"")}</div><div><span>In shop (DRIS)</span>${m.acceptedAt?fmtDate(m.acceptedAt.slice(0,10)+"T00:00"):"—"}</div><div><span>Days in shop</span>${m.acceptedAt?srDays(m):"—"}</div><div><span>Hour meter</span>${it?.hours??"—"}</div></div>
  <h2>Problem</h2><div>${(m.codes||[]).map(c=>`<b>${e(c)}</b> ${e(codeName(c))}`).join(" · ")}</div><div class="box">${e(m.desc||"")}</div>
  <h2>Tasks</h2><table><thead><tr><th>#</th><th>Task</th><th>Mechanic</th><th>Status</th><th class="r">Planned</th><th class="r">Actual hrs</th></tr></thead><tbody>${tasks}</tbody></table>
  ${labor?`<h2>Labor</h2><table><thead><tr><th>Date</th><th>Mechanic</th><th>Task</th><th>Note</th><th class="r">Hours</th></tr></thead><tbody>${labor}</tbody></table>`:""}
  ${parts?`<h2>Parts</h2><table><thead><tr><th>Part</th><th>Part #</th><th>Source</th><th>Status</th><th class="r">Qty</th><th class="r">Cost</th></tr></thead><tbody>${parts}</tbody></table>`:""}
  <h2>Cost</h2><table><tbody><tr><td>Labor (${+srLabor(m).toFixed(2)} hrs × ${money(laborRate())})</td><td class="r">${money(srLabor(m)*laborRate())}</td></tr><tr><td>Parts</td><td class="r">${money(partsCost(m))}</td></tr>${+m.cost?`<tr><td>Outside services</td><td class="r">${money(+m.cost)}</td></tr>`:""}<tr><td><b>Total</b></td><td class="r"><b>${money(srCost(m))}</b></td></tr></tbody></table>
  <h2>Final inspection</h2>${m.qc?`<div>Passed · ${e(m.qc.by||"")} · ${fmtDate(m.qc.at)}<div class="sub">${e(m.qc.note||"")}</div></div>`:`<div class="sig"><div><span></span>Inspected by</div><div><span></span>Date</div></div>`}
  </div>${autoPrint?`<script>window.onload=()=>setTimeout(()=>print(),300)<\/script>`:""}</body></html>`;
}

/* ---- Maintenance page: queues like a unit work queue ---- */
function vMaint(){
  const u = ui.maint, open = S.maintenance.filter(m=>m.status==="Open"), mine = myName();
  const Q = {
    open: open, mine: open.filter(m=>(m.tasks||[]).some(t=>t.owner===mine && t.status!=="Complete" && t.status!=="Cancelled")),
    submitted: open.filter(m=>m.stage==="Submitted"), parts: open.filter(m=>m.stage==="Awaiting parts"), qc: open.filter(m=>m.stage==="QC inspection"),
    deadline: open.filter(m=>m.opStatus==="Deadlined"), done: S.maintenance.filter(m=>m.status==="Completed"), all: S.maintenance};
  if (!Q[u.tab]) u.tab = "open";
  let list = Q[u.tab].filter(m=>(!u.sev || m.severity===u.sev) && (!u.type || m.type===u.type));
  if (u.q){ const q = u.q.toLowerCase(); list = list.filter(m=>[srNo(m), m.itemId, itemById(m.itemId)?.name, m.desc, m.tech, (m.codes||[]).join(" ")].join(" ").toLowerCase().includes(q)); }
  list = list.slice().sort((a,b)=> u.tab==="done"||u.tab==="all" ? (b.doneAt||b.date).localeCompare(a.doneAt||a.date) : ((a.opStatus==="Deadlined"?0:1)-(b.opStatus==="Deadlined"?0:1)) || (SEV_RANK[a.severity]??3)-(SEV_RANK[b.severity]??3) || srDays(b)-srDays(a));
  const inShop = open.filter(m=>m.stage!=="Submitted"), avgDays = inShop.length ? Math.round(inShop.reduce((s,m)=>s+srDays(m),0)/inShop.length) : 0;
  const yr = String(now().getFullYear()), ytd = S.maintenance.filter(m=>(m.doneAt||m.date).slice(0,4)===yr).reduce((s,m)=>s+srCost(m),0);
  const late = inShop.filter(m=>srDays(m)>srTarget()).length;
  return band("Maintenance","Service requests move from Submitted → Accepted into the shop → In work → QC inspection → Closed. Deadlined or accepted assets can't be rented.",`<button class="btn primary" data-act="new-maint">+ New service request</button>`) +
  tabs([["open","Open",Q.open.length],["mine","My tasks",Q.mine.length],["submitted","Submitted",Q.submitted.length],["parts","Awaiting parts",Q.parts.length],["qc","QC",Q.qc.length],["deadline","Deadlined",Q.deadline.length],["done","Closed",Q.done.length],["all","All",Q.all.length]],u.tab,"mtab") + `
  <div class="content">
    <div class="kpis">
      <div class="kpi"><span class="lbl">Open requests</span><span class="val num">${open.length}</span><span class="note">${Q.submitted.length} waiting to be accepted</span></div>
      <div class="kpi ${Q.deadline.length?"bad":""}"><span class="lbl">Deadlined</span><span class="val num">${Q.deadline.length}</span><span class="note">down, can't be used</span></div>
      <div class="kpi"><span class="lbl">Awaiting parts</span><span class="val num">${Q.parts.length}</span><span class="note">${srOpenParts({tasks:open.flatMap(m=>m.tasks||[])}).length} parts needed or on order</span></div>
      <div class="kpi"><span class="lbl">Ready for QC</span><span class="val num">${Q.qc.length}</span><span class="note">waiting on final inspection</span></div>
      <div class="kpi ${late?"bad":""}"><span class="lbl">Avg days in shop</span><span class="val num">${avgDays}</span><span class="note">${late?`${late} over the ${srTarget()}-day target`:`target ${srTarget()} days`}</span></div>
      <div class="kpi"><span class="lbl">Cost this year</span><span class="val num">${money0(ytd)}</span><span class="note">labor, parts and outside work</span></div>
    </div>
    <div class="toolbar" style="margin-bottom:10px;gap:8px;flex-wrap:wrap">
      <div class="search" style="max-width:300px"><span>${ico("search")}</span><input id="m-q" type="search" placeholder="SR #, asset, mechanic, code" value="${esc(u.q||"")}" data-input="mq" aria-label="Search service requests"></div>
      <select data-change="msev" aria-label="Severity"><option value="">All severities</option>${SEVERITY.map(([v])=>`<option ${u.sev===v?"selected":""}>${v}</option>`).join("")}</select>
      <select data-change="mtype" aria-label="Request type"><option value="">All request types</option>${REQ_TYPES.map(v=>`<option ${u.type===v?"selected":""}>${v}</option>`).join("")}</select>
    </div>
    <div class="tablewrap"><table><thead><tr><th>SR #</th><th>Asset</th><th>Type</th><th>Severity</th><th>Op status</th><th>Status</th><th>Tasks</th><th>Parts</th><th>Mechanic</th><th style="text-align:right">Days</th><th style="text-align:right">Cost</th><th>Docs</th><th></th></tr></thead>
    <tbody>${list.length?list.map(m=>{ const it = itemById(m.itemId), tk = (m.tasks||[]).filter(t=>t.status!=="Cancelled"), td = tk.filter(t=>t.status==="Complete").length, op = srOpenParts(m).length, d = srDays(m), lt = m.status==="Open" && m.acceptedAt && d>srTarget();
      return `<tr class="clickable" data-act="view-sr" data-id="${m.id}"><td class="num"><b>${srNo(m)}</b><div class="small muted">${fmtDate(m.date+"T00:00")}</div></td><td><button class="linkbtn" data-act="view-item" data-id="${esc(m.itemId)}">#${esc(m.itemId)}</button> ${esc(it?.name||"")}<div class="small muted trunc" title="${esc(m.desc)}">${(m.codes||[]).map(c=>`<b>${esc(c)}</b>`).join(" ")} ${esc(m.desc)}</div></td><td class="small">${esc(m.type)}</td><td>${sevChip(m.severity)}</td><td>${chipFor(m.opStatus||"Operational")}</td><td>${chipFor(m.stage)}</td>
      <td class="num">${tk.length?`${td}/${tk.length}`:"—"}</td><td class="num">${op?`<span style="color:var(--warn);font-weight:700">${op} open</span>`:srParts(m).length?"✓":"—"}</td><td class="small">${esc(m.tech||"")||'<span class="muted">Unassigned</span>'}</td>
      <td class="num" style="text-align:right">${m.acceptedAt?`<span ${lt?'style="color:var(--bad);font-weight:700"':""}>${d}</span>`:'<span class="muted">—</span>'}</td><td class="num" style="text-align:right">${srCost(m)?money(srCost(m)):"—"}</td><td>${docsBadge(m)||'<span class="muted">—</span>'}</td><td><button class="btn sm" data-act="view-sr" data-id="${m.id}">Open</button></td></tr>`; }).join("")
      : `<tr><td colspan="13"><div class="empty"><strong>Nothing in this queue</strong>Use <b>New service request</b> for repairs, inspections and preventive maintenance. Returns can also send an asset straight to the shop.</div></td></tr>`}</tbody></table></div>
  </div>`;
}

/* ---------------- Superintendent requests: browse what's available, request it, admins approve ---------------- */
const isSuper = () => me()?.role === "Superintendent";
const SUPER_ROUTES = ["request","my-requests","my-reports"];
const rqNo = r => "RQ-" + r.no;
const rqWhen = (d, t) => d ? fmtDate(d + "T" + (t || "00:00")) + (t ? " " + fmtTime(t) : "") : "";
function fmtTime(t){ const [h,m] = t.split(":").map(Number); return `${((h+11)%12)+1}:${String(m).padStart(2,"0")} ${h<12?"AM":"PM"}`; }
ui.req = {tab:"assets", q:"", cat:"", cart:[]};
ui.rqa = {tab:"pending"};
function reqCartKey(){ return KEY + "-cart-" + (me()?.id||""); }
function loadCart(){ try { const c = JSON.parse(localStorage.getItem(reqCartKey())||"[]"); ui.req.cart = Array.isArray(c) ? c : []; } catch(e){ ui.req.cart = []; } }
function saveCart(){ try { localStorage.setItem(reqCartKey(), JSON.stringify(ui.req.cart)); } catch(e){} }
const superAssets = () => S.items.filter(i=>i.status!=="Retired" && !i.shopOnly);
const superStock = () => S.stock.filter(s=>!s.retired && (+s.qty||0) - stockOut(s.id) > 0);
const superInv = () => S.inventory.filter(i=>!i.retired);
const pendingRequests = () => (S.requests||[]).filter(r=>r.status==="Pending");
function reqRef(l){ return l.kind==="asset" ? itemById(l.id) : l.kind==="stock" ? S.stock.find(s=>s.id===l.id) : S.inventory.find(i=>i.id===l.id); }
function reqLineName(l){ const r = reqRef(l); return l.kind==="asset" ? `#${l.id} ${r?.name||"(removed)"}` : (r?.name||"(removed)"); }
function reqLineUnit(l){ const r = reqRef(l); return l.kind==="asset" ? "" : (r?.unit||"ea"); }
function reqLineAvail(l){ const r = reqRef(l); if (!r) return {ok:false, txt:"No longer in the system"};
  if (l.kind==="asset"){ const st = itemStatus(r); return {ok:st==="Available", txt:st}; }
  if (l.kind==="stock"){ const a = (+r.qty||0) - stockOut(r.id); return {ok:a>=l.qty, txt:`${a} ${r.unit||"ea"} available`}; }
  return {ok:(+r.qty||0)>=l.qty, txt:`${r.qty} ${r.unit||"ea"} on hand`}; }
/* rate text for any catalog item */
function rateParts(x){ return [["day","Day"],["week","Week"],["month","Month"]].filter(([k])=>+x[k]>0).map(([k,l])=>[l, +x[k]]); }
function rateHtml(x, kind){
  if (kind==="inv") return canSell(x) ? `<div class="rq-rates"><span><b>${money(+x.price)}</b> per ${esc(x.unit||"ea")}</span></div>` : `<div class="rq-rates muted"><span>Check-out only · no charge</span></div>`;
  const r = rateParts(x); const per = kind==="stock" ? ` <i>per ${esc(x.unit||"ea")}</i>` : "";
  return r.length ? `<div class="rq-rates">${r.map(([l,v])=>`<span><small>${l}</small><b>${money0(v)}</b></span>`).join("")}${per?`<span class="rq-per">${per}</span>`:""}</div>` : `<div class="rq-rates muted"><span>${kind==="stock"?"No charge":"Rate on request"}</span></div>`;
}
function reqEstimate(){
  const f = $("#rqform"); if (!f) return; const a = $("#rq-need").value, at = $("#rq-need-t").value||"08:00", b = $("#rq-until").value, bt = $("#rq-until-t").value||"08:00";
  let total = 0, ok = !!(a && b && (b+"T"+bt) > (a+"T"+at));
  ui.req.cart.forEach((c,i)=>{ const el = $("#rq-est-"+i); if (!el) return; const r = reqRef(c); let v = null;
    if (c.kind==="inv"){ if (r && canSell(r)) v = (+r.price||0) * (+c.qty||0); }
    else if (ok && r && rateParts(r).length){ const p = priceLine({day:+r.day>0?+r.day:null, week:+r.week>0?+r.week:null, month:+r.month>0?+r.month:null, qty:1}, a+"T"+at, b+"T"+bt); v = p.total * (c.kind==="stock" ? (+c.qty||1) : 1); }
    if (v!=null){ total += v; el.textContent = money(v); } else el.textContent = c.kind==="inv" ? "No charge" : ok ? "—" : "Add dates"; });
  const t = $("#rq-est-total"); if (t) t.textContent = money(total);
}
const inCart = (kind,id) => ui.req.cart.find(c=>c.kind===kind && c.id===id);
const nmSort = (a,b) => az(String(a.name||"").trim().replace(/^[^a-z0-9]+/i,""), String(b.name||"").trim().replace(/^[^a-z0-9]+/i,""));

function vRequestCatalog(){
  const U = ui.req; if (!U.loaded){ loadCart(); U.loaded = me()?.id; } else if (U.loaded !== me()?.id){ loadCart(); U.loaded = me()?.id; }
  const A = superAssets(), K = superStock(), I = superInv();
  const src = U.tab==="stock" ? K : U.tab==="inv" ? I : A;
  let list = src.slice();
  if (U.cat) list = list.filter(x=>x.category===U.cat || groupOf(x.category)?.parent===U.cat);
  if (U.q){ const q = U.q.toLowerCase(); list = list.filter(x=>[x.id, x.name, x.category, x.make, x.model, x.sku].join(" ").toLowerCase().includes(q)); }
  list.sort((a,b)=> U.tab==="assets" ? az(a.id,b.id) : nmSort(a,b));
  const kind = U.tab==="stock" ? "stock" : U.tab==="inv" ? "inv" : "asset";
  const cats = [...new Set(src.map(x=>x.category).filter(Boolean))];
  const card = x => { const c = inCart(kind, x.id), photo = kind==="asset" ? primaryPhoto(x)?.data : x.photo;
    const sub = kind==="asset" ? `#${esc(x.id)}${x.category?` · ${esc(x.category)}`:""}` : `${x.category?esc(x.category):""}`;
    const qtyTxt = kind==="asset" ? chipFor(itemStatus(x)) : kind==="stock" ? `<span class="small"><b class="num">${(+x.qty||0)-stockOut(x.id)}</b> ${esc(x.unit||"ea")} available</span>` : `<span class="small"><b class="num">${x.qty}</b> ${esc(x.unit||"ea")} on hand</span>${invStatus(x)!=="In stock"?` <span class="chip ${x.qty<=0?"overdue":"maint"}">${invStatus(x)}</span>`:""}`;
    return `<div class="rq-card ${c?"in":""}"><div class="rq-ph">${photo?`<img src="${photo}" alt="">`:ico(kind==="asset"?"items":kind==="stock"?"stock":"inventory",1.4)}</div>
      <div class="rq-body"><b class="rq-name">${esc(x.name)}</b><span class="small muted">${sub}</span><div class="rq-meta">${qtyTxt}</div>${rateHtml(x, kind)}</div>
      <div class="rq-act">${c ? (kind==="asset" ? `<button class="btn sm" data-act="rq-remove" data-kind="${kind}" data-id="${esc(x.id)}">✓ Added</button>` :
        `<div class="rq-step"><button class="btn sm" data-act="rq-dec" data-kind="${kind}" data-id="${esc(x.id)}" aria-label="Less">−</button><span class="num">${c.qty}</span><button class="btn sm" data-act="rq-inc" data-kind="${kind}" data-id="${esc(x.id)}" aria-label="More">+</button></div>`)
        : (kind==="asset" && itemStatus(x)!=="Available") ? '<span class="small muted">Not available now</span>' : `<button class="btn sm primary" data-act="rq-add" data-kind="${kind}" data-id="${esc(x.id)}">+ Request</button>`}</div></div>`; };
  const n = U.cart.length;
  return band("Equipment catalog", can("request") ? "The whole fleet, small equipment and inventory, with rates. Add what your job needs and send the request to the Equipment Manager." : "The whole fleet, small equipment and inventory, with rates and what is available now.", `<button class="btn" data-act="goto" data-to="my-requests">My requests</button><button class="btn primary" data-act="rq-review" ${n?"":"disabled"}>Review request${n?` (${n})`:""}</button>`) +
  tabs([["assets","Equipment",A.length],["stock","Small equipment",K.length],["inv","Inventory",I.length]], U.tab, "rqtab") + `
  <div class="content ${n?"rq-pad":""}">
    <div class="toolbar" style="margin-bottom:12px;gap:8px;flex-wrap:wrap">
      <div class="search" style="max-width:340px">${ico("search")}<input id="rq-q" type="search" placeholder="${kind==="asset"?"Asset #, name, make or model":"Search by name"}" value="${esc(U.q)}" data-input="rqq" aria-label="Search"></div>
      <select data-change="rqcat" aria-label="Group"><option value="">All groups</option>${orderedGroups().filter(g=>cats.includes(g.name) || cats.some(c=>groupOf(c)?.parent===g.name)).map(g=>`<option value="${esc(g.name)}" ${U.cat===g.name?"selected":""}>${g.parent?"   › ":""}${esc(g.name)}</option>`).join("")}</select>
      <span class="small muted">${list.length} ${kind==="asset"?"pieces of equipment":kind==="stock"?"small equipment items":"inventory items"}</span>
    </div>
    <div class="rq-grid">${list.map(card).join("") || `<div class="empty" style="grid-column:1/-1"><strong>Nothing matches</strong>Try another search or group.</div>`}</div>
  </div>
  ${n?`<div class="rq-bar"><span><b>${n}</b> item${n>1?"s":""} in your request</span><button class="btn" data-act="rq-clear">Clear</button><button class="btn primary" data-act="rq-review">Review & send</button></div>`:""}`;
}
function reqReview(){
  const cart = ui.req.cart; if (!cart.length) return;
  const d1 = new Date(now().getTime()+DAY), d2 = new Date(now().getTime()+(+S.settings.cycle||28)*DAY), hasEq = cart.some(c=>c.kind!=="inv");
  const lastLoc = (S.requests||[]).filter(r=>r.by===me()?.id).slice(-1)[0]?.locationId;
  openModal(`<header><h2>Review your request</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="rqform" class="form">
    ${cloudJobSelect()}
    <div class="field"><label for="rq-need">Needed on site by</label><div class="dt-pair"><input id="rq-need" type="date" name="needBy" value="${toLocalInput(d1).slice(0,10)}"><input id="rq-need-t" type="time" name="needTime" value="08:00" aria-label="Time needed" step="900"></div></div>
    <div class="field"><label for="rq-until">${hasEq?"Return by (equipment)":"Return by"}</label><div class="dt-pair"><input id="rq-until" type="date" name="returnBy" value="${hasEq?toLocalInput(d2).slice(0,10):""}"><input id="rq-until-t" type="time" name="returnTime" value="${hasEq?"08:00":""}" aria-label="Return time" step="900"></div><span class="hint">${hasEq?"Best guess is fine. The admin can adjust it.":"Leave blank for supplies that won't come back"}</span></div>
    ${cloudDeliveryFields()}<div class="field full"><label for="rq-notes">Notes</label><textarea id="rq-notes" name="notes" placeholder="Anything the equipment team should know"></textarea></div>
  </form>
  <div class="subhead"><span>Items (${cart.length})</span></div>
  <div class="rq-lines">${cart.map((c,i)=>`<div class="rq-line"><span class="chip neutral">${c.kind==="asset"?"Equipment":c.kind==="stock"?"Small equip.":"Inventory"}</span><b>${esc(reqLineName(c))}</b>
    ${c.kind==="asset"?'<span class="small muted">1</span>':`<input type="number" min="0.1" step="any" value="${c.qty}" data-rqi="${i}" aria-label="Quantity" class="sm-input num" style="width:80px"> <span class="small muted">${esc(reqLineUnit(c))}</span>`}
    <span class="rq-est num" id="rq-est-${i}"></span>
    <button class="del" data-act="rq-line-del" data-i="${i}" aria-label="Remove">×</button>
    <div class="rq-line-rate">${rateHtml(reqRef(c)||{}, c.kind)}</div></div>`).join("")}
    <div class="rq-line rq-total"><b>Estimated total</b><span class="rq-est num" id="rq-est-total"></span><span class="small muted" style="flex-basis:100%">Based on the dates above and the cheapest mix of day, week and month rates. Delivery, fuel and tax are extra. The final bill comes from the order.</span></div></div>
  </div>
  <footer><button class="btn" data-act="close">Keep browsing</button><button class="btn primary" data-act="rq-send">Send request</button></footer>`, "narrow");
  reqEstimate();
}
function reqSend(){
  const fd = Object.fromEntries(new FormData($("#rqform"))), u = me();
  if (!fd.locationId){ toast("Pick the job"); return; }
  if (fd.returnBy && fd.needBy && (fd.returnBy+"T"+(fd.returnTime||"23:59")) <= (fd.needBy+"T"+(fd.needTime||"00:00"))){ toast("Return date is before the date it's needed"); return; }
  const cust = S.customers.find(c=>!c.deleted && c.name.trim().toLowerCase()===String(u?.name||"").trim().toLowerCase());
  S.settings.nextRQ ||= 1001;
  const r = {id:uid("rq"), no:S.settings.nextRQ++, by:u?.id, byName:u?.name||"", customerId:cust?.id||"", locationId:fd.locationId, needBy:fd.needBy, needTime:fd.needTime||"", returnBy:fd.returnBy, returnTime:fd.returnBy?(fd.returnTime||""):"", notes:fd.notes.trim(),
    lines:ui.req.cart.map(c=>({...c})), status:"Pending", createdAt:toLocalInput(now()), log:[{at:toLocalInput(now()), by:u?.name||"", text:"Request sent"}]};
  (S.requests ||= []).push(r); ui.req.cart = []; saveCart(); save();
  const admins = S.users.filter(x=>x.active!==false && x.role==="Admin" && x.email).map(x=>x.email);
  const body = `${rqNo(r)} from ${r.byName}\nJob: ${locLabel(locById(r.locationId))}\nNeeded by: ${rqWhen(r.needBy, r.needTime)}${r.returnBy?`\nReturn by: ${rqWhen(r.returnBy, r.returnTime)}`:""}\n\n${r.lines.map(l=>`- ${reqLineName(l)}${l.kind==="asset"?"":` × ${l.qty} ${reqLineUnit(l)}`}`).join("\n")}${r.notes?`\n\nNotes: ${r.notes}`:""}\n\nOpen the equipment app → Requests to approve it.`;
  openModal(`<header><h2>${rqNo(r)} sent</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><p>Your request is in the <b>Requests</b> list for the equipment admins. You'll see their answer under <b>My requests</b>.</p>
  ${admins.length?`<p class="small muted">Want them to know right away? Send a quick email too.</p><a class="btn" href="mailto:${admins.map(encodeURIComponent).join(",")}?subject=${encodeURIComponent(`Equipment request ${rqNo(r)} · ${locLabel(locById(r.locationId))}`)}&body=${encodeURIComponent(body)}">Email the admins</a>`:""}</div>
  <footer><button class="btn" data-act="close">Request more</button><button class="btn primary" data-act="goto" data-to="my-requests">See my requests</button></footer>`, "narrow");
  rerender();
}
function reqCardHtml(r, admin){
  const L = locById(r.locationId), problems = r.status==="Pending" ? r.lines.filter(l=>!reqLineAvail(l).ok) : [];
  return `<div class="rq-req">
    <div class="rq-req-h"><b class="num">${rqNo(r)}</b>${chipFor(r.status)}${admin?`<span><b>${esc(r.byName)}</b></span>`:""}<span class="muted small">${fmtDate(r.createdAt)}</span>${r.orderNo?`<button class="linkbtn small" data-act="view-order" data-no="${r.orderNo}">Order #${r.orderNo}</button>`:""}</div>
    <div class="rq-req-meta"><span>${ico("locations")} ${esc(locLabel(L))}</span><span>Needed <b>${rqWhen(r.needBy, r.needTime)}</b></span>${r.returnBy?`<span>Return <b>${rqWhen(r.returnBy, r.returnTime)}</b></span>`:""}</div>
    <ul class="rq-req-lines">${r.lines.map(l=>{ const a = reqLineAvail(l); return `<li><span>${esc(reqLineName(l))}${l.kind==="asset"?"":` <b>× ${l.qty}</b> ${esc(reqLineUnit(l))}`}</span>${admin&&r.status==="Pending"?`<span class="small ${a.ok?"muted":""}" ${a.ok?"":'style="color:var(--bad);font-weight:700"'}>${esc(a.txt)}</span>`:""}</li>`; }).join("")}</ul>
    ${cloudReqExtra(r, admin)}${r.notes?`<div class="rq-note">${esc(r.notes)}</div>`:""}
    ${r.decisionNote?`<div class="rq-note admin"><b>${esc(r.decidedBy||"Admin")}:</b> ${esc(r.decisionNote)}</div>`:""}
    ${admin&&problems.length?`<div class="small" style="color:var(--bad)">${problems.length} item${problems.length>1?"s are":" is"} no longer available. You can still approve the rest.</div>`:""}
    <div class="rq-req-f">${admin && r.status==="Pending" ? `<button class="btn" data-act="rq-decline" data-id="${r.id}">Decline</button><button class="btn" data-act="rq-approve" data-id="${r.id}" title="Approve without creating a rental order (for check-outs or supplies)">Approve only</button><button class="btn primary" data-act="rq-order" data-id="${r.id}">Approve & create order</button>`
      : !admin && r.status==="Pending" ? `<button class="btn" data-act="rq-cancel" data-id="${r.id}">Cancel request</button>` : ""}</div>
  </div>`;
}
function vMyRequests(){
  const mine = (S.requests||[]).filter(r=>r.by===me()?.id).sort((a,b)=>b.no-a.no);
  return band("My requests", "Requests you've sent and what the equipment team decided.", `<button class="btn primary" data-act="goto" data-to="request">+ New request</button>`) + `
  <div class="content"><div class="rq-list">${mine.map(r=>reqCardHtml(r,false)).join("") || `<div class="empty"><strong>No requests yet</strong>Go to <b>Request</b>, add equipment or inventory, and send it.</div>`}</div></div>`;
}
function vRequestsAdmin(){
  const all = (S.requests||[]).slice().sort((a,b)=>b.no-a.no), t = ui.rqa.tab;
  const list = all.filter(r=> t==="all" || (t==="pending" ? r.status==="Pending" : t==="approved" ? r.status==="Approved" : r.status==="Declined"||r.status==="Cancelled"));
  return band("Requests", "Equipment and inventory requests from superintendents. Approve one to turn it into a booked rental order, or decline it with a note.", "") +
  tabs([["pending","Pending",all.filter(r=>r.status==="Pending").length],["approved","Approved",all.filter(r=>r.status==="Approved").length],["declined","Declined & cancelled",all.filter(r=>r.status==="Declined"||r.status==="Cancelled").length],["all","All",all.length]], t, "rqatab") + `
  <div class="content"><div class="rq-list">${list.map(r=>reqCardHtml(r,true)).join("") || `<div class="empty"><strong>${t==="pending"?"No requests waiting":"Nothing here"}</strong>${t==="pending"?"When a superintendent sends a request, it shows up here and on the bell at the top.":""}</div>`}</div></div>`;
}
function reqDecide(r, status, note, orderNo){
  r.status = status; r.decidedBy = myName(); r.decidedAt = toLocalInput(now()); if (note) r.decisionNote = note; if (orderNo) r.orderNo = orderNo;
  (r.log ||= []).push({at:r.decidedAt, by:myName(), text:`${status}${orderNo?` as order #${orderNo}`:""}${note?`: ${note}`:""}`});
}
function reqNoteBox(r, mode){
  openModal(`<header><h2>${mode==="decline"?"Decline":"Approve"} ${rqNo(r)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><div class="field"><label for="rq-dnote">Note to ${esc(r.byName||"the superintendent")}${mode==="decline"?"":" (optional)"}</label><textarea id="rq-dnote" placeholder="${mode==="decline"?"Why, or what you can offer instead":"Delivery time, what to expect"}"></textarea></div></div>
  <footer><button class="btn" data-act="close">Cancel</button><button class="btn ${mode==="decline"?"danger":"primary"}" data-act="rq-${mode}-save" data-id="${r.id}">${mode==="decline"?"Decline request":"Approve"}</button></footer>`, "narrow");
}
function reqToOrder(r){ openOrderForm(reqDraft(r)); if (!r.customerId) toast(`No customer named ${r.byName} yet. Pick the customer for this order`); }
function reqDraft(r){
  const d = newOrderDraft();
  d.status = "Booked"; d.locationId = r.locationId; d.customerId = r.customerId && custById(r.customerId) ? r.customerId : "";
  d.out = (r.needBy||toLocalInput(now()).slice(0,10)) + "T" + (r.needTime||"08:00"); if (r.returnBy) d.ret = r.returnBy + "T" + (r.returnTime||"08:00"); else d.ret = toLocalInput(new Date(parseD(d.out).getTime()+(+S.settings.cycle||28)*DAY));
  d.lines = r.lines.filter(l=>l.kind==="asset" && itemById(l.id) && itemById(l.id).status!=="Retired").map(l=>lineFrom(l.id));
  d.sales = r.lines.filter(l=>l.kind==="inv" && canSell(reqRef(l))).map(l=>({invId:l.id, qty:l.qty, price:+reqRef(l).price||0}));
  const other = r.lines.filter(l=>l.kind==="stock" || (l.kind==="inv" && !canSell(reqRef(l))));
  d.description = `From ${rqNo(r)} (${r.byName})${other.length?`. Check out separately: ${other.map(l=>`${reqLineName(l)} × ${l.qty}`).join(", ")}`:""}${r.notes?`. Notes: ${r.notes}`:""}`;
  d.fromRequest = r.id;
  return d;
}
function reqAction(a, ds){
  const r = ds.id ? (S.requests||[]).find(x=>x.id===ds.id) : null;
  const cart = ui.req.cart;
  switch(a){
    case "rq-add": { if (!inCart(ds.kind, ds.id)) cart.push({kind:ds.kind, id:ds.id, qty:1}); saveCart(); rerender(); return true; }
    case "rq-remove": ui.req.cart = cart.filter(c=>!(c.kind===ds.kind && c.id===ds.id)); saveCart(); rerender(); return true;
    case "rq-inc": { const c = inCart(ds.kind, ds.id); if (c) c.qty = +(c.qty+1).toFixed(2); saveCart(); rerender(); return true; }
    case "rq-dec": { const c = inCart(ds.kind, ds.id); if (c){ c.qty = +(c.qty-1).toFixed(2); if (c.qty<=0) ui.req.cart = cart.filter(x=>x!==c); } saveCart(); rerender(); return true; }
    case "rq-clear": ui.req.cart = []; saveCart(); rerender(); return true;
    case "rq-review": reqReview(); return true;
    case "rq-line-del": cart.splice(+ds.i,1); saveCart(); if (cart.length) reqReview(); else closeModal(); rerender(); return true;
    case "rq-send": reqSend(); return true;
    case "rq-cancel": reqDecide(r, "Cancelled", ""); r.decidedBy = ""; save(); rerender(); toast(`${rqNo(r)} cancelled`); return true;
    case "rq-decline": reqNoteBox(r, "decline"); return true;
    case "rq-approve": reqNoteBox(r, "approve"); return true;
    case "rq-decline-save": { const n = $("#rq-dnote").value.trim(); if (!n){ toast("Add a short note so they know why"); return true; } reqDecide(r, "Declined", n); save(); closeModal(); rerender(); updateBell(); toast(`${rqNo(r)} declined`); return true; }
    case "rq-approve-save": reqDecide(r, "Approved", $("#rq-dnote").value.trim()); save(); closeModal(); rerender(); updateBell(); toast(`${rqNo(r)} approved`); return true;
    case "rq-order": reqToOrder(r); return true;
  }
  return false;
}


/* ---- Superintendent reports: only their own orders ---- */
ui.myrep = {tab:"onrent", period:"all", loc:""};
function superCustomer(u){ u = u || me(); return (u?.customerId && custById(u.customerId)) || S.customers.find(c=>!c.deleted && c.name.trim().toLowerCase()===String(u?.name||"").trim().toLowerCase()); }
function myOrders(){ const u = me(), c = superCustomer(u), rq = new Set((S.requests||[]).filter(r=>r.by===u?.id).map(r=>r.no));
  return S.orders.filter(o=>(c && o.customerId===c.id) || (o.requestNo && rq.has(o.requestNo))); }
function myRepData(){
  const U = ui.myrep, n = now(), all = myOrders();
  const start = U.period==="month" ? new Date(n.getFullYear(), n.getMonth(), 1) : U.period==="90" ? new Date(n.getTime()-90*DAY) : U.period==="ytd" ? new Date(n.getFullYear(),0,1) : null;
  let orders = all.filter(o=>o.status!=="Cancelled");
  if (U.loc) orders = orders.filter(o=>o.locationId===U.loc);
  const inPeriod = o => !start || parseD(o.status==="Rented Out" ? nowStr() : (o.returnedAt||o.ret)) >= start;
  const per = orders.filter(inPeriod);
  const active = orders.filter(o=>o.status==="Rented Out");
  const onrent = active.flatMap(o=>o.lines.map(l=>{ const it = itemById(l.itemId), days = Math.max(0, Math.floor((n-parseD(o.out))/DAY)), acc = priceLine(l, o.out, nowStr()).total;
    return {no:o.no, id:l.itemId, name:it?.name||"", job:locLabel(locById(o.locationId)), out:o.out, due:o.ret, days, late:isOverdue(o), month:+l.month||0, accrued:acc}; }));
  const hist = per.slice().sort((a,b)=>b.no-a.no).map(o=>{ const t = orderTotals(o); return {o, job:locLabel(locById(o.locationId)), items:o.lines.map(l=>"#"+l.itemId).join(", ")+((o.sales||[]).length?` + ${(o.sales||[]).length} supply line${o.sales.length>1?"s":""}`:""), rental:t.rental, supplies:t.sales||0, other:t.delivery+t.charges+(+o.adjustment||0), total:t.total}; });
  const byJob = [...hist.reduce((m,h)=>{ const k = h.o.locationId, g = m.get(k) || {job:h.job, orders:0, items:0, rental:0, supplies:0, other:0, total:0, active:0}; g.orders++; g.items += h.o.lines.length; g.rental += h.rental; g.supplies += h.supplies; g.other += h.other; g.total += h.total; if (h.o.status==="Rented Out") g.active++; m.set(k,g); return m; }, new Map()).values()].sort((a,b)=>b.total-a.total);
  const sheets = [
    {name:"On rent now", cols:[["Order #","n",9],["Asset #","s",9],["Asset","s",28],["Job","s",30],["Rented out","d",16],["Due back","d",16],["Days out","n",9],["Overdue","s",9],["Month rate","$",12],["Accrued to date","$",14]], rows:onrent.map(r=>[r.no,r.id,r.name,r.job,r.out,r.due,r.days,r.late?"Yes":"",r.month,r.accrued]), totals:true},
    {name:"Order history", cols:[["Order #","n",9],["Job","s",30],["Status","s",12],["Rented out","d",16],["Return","d",16],["Items","s",30],["Rental","$",12],["Supplies","$",12],["Delivery & other","$",14],["Total","$",12]], rows:hist.map(h=>[h.o.no,h.job,orderStatus(h.o),h.o.out,h.o.returnedAt||h.o.ret,h.items,h.rental,h.supplies,h.other,h.total]), totals:true},
    {name:"Cost by job", cols:[["Job","s",32],["Orders","n",9],["Active","n",9],["Items","n",9],["Rental","$",12],["Supplies","$",12],["Delivery & other","$",14],["Total","$",12]], rows:byJob.map(g=>[g.job,g.orders,g.active,g.items,g.rental,g.supplies,g.other,g.total]), totals:true}];
  return {all, orders, per, active, onrent, hist, byJob, sheets};
}
function vMyReports(){
  const U = ui.myrep, c = superCustomer(), D = myRepData();
  const jobs = [...new Map(D.all.map(o=>[o.locationId, locLabel(locById(o.locationId))])).entries()].sort((a,b)=>az(a[1],b[1]));
  const per = [["month","This month"],["90","Last 90 days"],["ytd","Year to date"],["all","All time"]];
  const tableHtml = sh => `<div class="tablewrap"><table><thead><tr>${sh.cols.map(([h,t])=>`<th ${t==="$"||t==="n"?'style="text-align:right"':""}>${h}</th>`).join("")}</tr></thead><tbody>${sh.rows.length?sh.rows.map(r=>`<tr>${r.map((v,i)=>{ const t = sh.cols[i][1], h = sh.cols[i][0];
      const out = v==null||v==="" ? '<span class="muted">—</span>' : t==="$" ? money(v) : t==="d" ? fmtDate(v) : h==="Order #" ? `<button class="linkbtn" data-act="my-order" data-no="${v}">#${v}</button>` : h==="Overdue" ? chipFor("Overdue") : h==="Status" ? chipFor(v) : esc(String(v));
      return `<td ${t==="$"||t==="n"?'class="num" style="text-align:right"':""}>${out}</td>`; }).join("")}</tr>`).join("")
    + (sh.totals?`<tr class="tot-row">${sh.cols.map(([h,t],i)=>`<td ${t==="$"||t==="n"?'class="num" style="text-align:right"':""}>${i===0?"<b>Total</b>":t==="$"?`<b>${money(sh.rows.reduce((s,r)=>s+(+r[i]||0),0))}</b>`:t==="n"&&!/^(Order #|Days)/.test(h)?`<b>${sh.rows.reduce((s,r)=>s+(+r[i]||0),0)}</b>`:""}</td>`).join("")}</tr>`:"")
    : `<tr><td colspan="${sh.cols.length}"><div class="empty">Nothing to show for these filters.</div></td></tr>`}</tbody></table></div>`;
  if (U.tab==="avail") return band("My reports", AVAIL_SUB, availBtns()) + tabs([["onrent","On rent now",D.onrent.length],["history","Order history",D.hist.length],["jobs","Cost by job",D.byJob.length],["avail","Available"]], U.tab, "mytab") + `<div class="content">${availBody()}</div>`;
  const idx = {onrent:0, history:1, jobs:2}[U.tab] ?? 0;
  const spend = D.hist.reduce((s,h)=>s+h.total,0);
  return band("My reports", c ? `Orders for <b>${esc(c.name)}</b>${c.position?` · ${esc(c.position)}`:""}. Only your own orders are included.` : "Your sign-in isn't linked to a customer yet, so only orders from your requests show. Ask an admin to link you on the Users page.",
    `<button class="btn" data-act="my-csv">${ico("download")} CSV</button><button class="btn primary" data-act="my-xlsx">${ico("download")} Excel (all 3)</button>`) +
  tabs([["onrent","On rent now",D.onrent.length],["history","Order history",D.hist.length],["jobs","Cost by job",D.byJob.length],["avail","Available"]], U.tab, "mytab") + `
  <div class="content">
    <div class="kpis">
      <div class="kpi"><span class="lbl">Active orders</span><span class="val num">${D.active.length}</span><span class="note">${D.onrent.length} asset${D.onrent.length===1?"":"s"} on rent</span></div>
      <div class="kpi ${D.onrent.some(r=>r.late)?"bad":""}"><span class="lbl">Overdue</span><span class="val num">${D.active.filter(isOverdue).length}</span><span class="note">past the return date</span></div>
      <div class="kpi"><span class="lbl">Accrued on rent now</span><span class="val num">${money0(D.onrent.reduce((s,r)=>s+r.accrued,0))}</span><span class="note">rent-out to today</span></div>
      <div class="kpi"><span class="lbl">Billed ${(per.find(x=>x[0]===U.period)||[,""])[1].toLowerCase()}</span><span class="val num">${money0(spend)}</span><span class="note">${D.hist.length} order${D.hist.length===1?"":"s"}</span></div>
    </div>
    <div class="toolbar" style="margin-bottom:10px;gap:8px;flex-wrap:wrap">
      <select data-change="myloc" aria-label="Job"><option value="">All my jobs</option>${jobs.map(([id,l])=>`<option value="${id}" ${U.loc===id?"selected":""}>${esc(l)}</option>`).join("")}</select>
      ${U.tab==="onrent"?"":`<select data-change="myper" aria-label="Period">${per.map(([k,l])=>`<option value="${k}" ${U.period===k?"selected":""}>${l}</option>`).join("")}</select>`}
    </div>
    ${tableHtml(D.sheets[idx])}
  </div>`;
}
function myOrderView(no){
  const o = myOrders().find(x=>x.no===+no); if (!o) return; const t = orderTotals(o);
  openModal(`<header><h2>Order #${o.no}</h2>${chipFor(orderStatus(o))}<button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><div class="sr-grid"><div><span>Job</span><b>${esc(locLabel(locById(o.locationId)))}</b></div><div><span>Rented out</span><b>${fmtDate(o.out)}</b></div><div><span>${o.status==="Returned"?"Returned":"Return date"}</span><b>${fmtDate(o.returnedAt||o.ret)}</b></div><div><span>Total</span><b class="num">${money(t.total)}</b></div></div>
  <div class="subhead"><span>Equipment</span></div><div class="rq-lines">${o.lines.map(l=>{ const p = priceLine(l, o.out, o.status==="Rented Out"?o.ret:(o.returnedAt||o.ret)); return `<div class="rq-line"><b>#${esc(l.itemId)} ${esc(itemById(l.itemId)?.name||"")}</b><span class="small muted">${esc(p.label||"")}</span><span class="num">${money(p.total)}</span></div>`; }).join("")||'<div class="small muted">No equipment.</div>'}</div>
  ${(o.sales||[]).length?`<div class="subhead"><span>Supplies</span></div><div class="rq-lines">${o.sales.map(x=>{ const it = S.inventory.find(i=>i.id===x.invId); return `<div class="rq-line"><b>${esc(it?.name||"Item")}</b><span class="small muted">${x.qty} ${esc(it?.unit||"")} × ${money(x.price)}</span><span class="num">${money(x.qty*x.price)}</span></div>`; }).join("")}</div>`:""}
  ${o.description?`<div class="sr-summary">${esc(o.description)}</div>`:""}</div>
  <footer><button class="btn" data-act="close">Close</button></footer>`, "narrow");
}

/* ---------------- Locations / Customers ---------------- */
function locStats(ids){
  const act = S.orders.filter(o=>ids.includes(o.locationId) && o.status==="Rented Out");
  return {act, itemsOn:act.reduce((s,o)=>s+o.lines.length,0),
    billed:S.orders.filter(o=>ids.includes(o.locationId) && o.status!=="Cancelled").reduce((s,o)=>s+orderTotals(o).total,0),
    parked:S.items.filter(it=>it.status!=="Retired" && ids.includes(it.locationId) && ["Available","In Maintenance"].includes(itemStatus(it))).length,
    over:act.filter(isOverdue).length};
}
function vLocations(){
  const open = ui.locs.open && locById(ui.locs.open) && locKids(ui.locs.open).length ? ui.locs.open : null;
  if (!open) ui.locs.open = null;
  const base = open ? [locById(open), ...locKids(open)] : S.locations.filter(l=>!locParent(l));
  const rows = base.map(l=>{ const self = open && l.id===open, kids = self ? 0 : locKids(l.id).length;
    return {l, kids, self, ...locStats(kids ? locFamily(l.id) : [l.id])}; })
    .sort((a,b)=> (b.self-a.self) || ((b.itemsOn+b.parked)-(a.itemsOn+a.parked)) || az(locLabel(a.l),locLabel(b.l)));
  const pl = open ? locById(open) : null, nk = open ? locKids(open).length : 0;
  const head = open
    ? band(locLabel(pl), `<button class="linkbtn" data-act="loc-close" data-id="">All locations</button>${locPath(pl).map(a=>` › <button class="linkbtn" data-act="loc-close" data-id="${a.id}">${esc(a.name)}</button>`).join("")} › ${esc(pl.name)} · ${nk} sub-location${nk>1?"s":""}${pl.address?" · "+esc(pl.address):""}`, `<button class="btn" data-act="edit-loc" data-id="${pl.id}">Edit ${esc(pl.name)}</button><button class="btn primary" data-act="new-loc" data-parent="${pl.id}">+ Add sub-location</button>`)
    : band("Locations","Job sites and the BHESC Shop Yard. Assets are kept at the shop yard unless they're out on a rental. Click a location with sub-locations to open it.",`<button class="btn primary" data-act="new-loc">+ Add location</button>`);
  return head + `
  <div class="content"><div class="tablewrap"><table><thead><tr><th>Job #</th><th>Name</th><th>Address</th><th>Site contact</th><th style="text-align:right">Kept here</th><th style="text-align:right">On rent here</th><th>Overdue</th><th style="text-align:right">Total billed</th><th></th></tr></thead>
  <tbody>${rows.map(({l,kids,self,itemsOn,billed,parked,over})=>`<tr class="${kids?"clickable":""} ${self?"own-row":""}" ${kids?`data-act="loc-open" data-id="${l.id}"`:""}><td class="num"><b>${esc(l.code)||'<span class="muted">—</span>'}</b></td><td><span class="grp-name">${kids?ico("locations"):""}${esc(l.name)}</span>${l.isYard?' <span class="chip available">Default</span>':""}${kids?`<div class="small grp-kids">${kids} sub-location${kids>1?"s":""} ›</div>`:""}${self?'<div class="small muted">This location itself</div>':""}</td><td class="muted">${esc(l.address)||"—"}</td><td>${esc(locContact(l))||"—"}</td><td class="num" style="text-align:right">${parked?`<button class="linkbtn" data-act="loc-assets" data-id="${l.id}"${self?' data-exact="1"':""}>${parked}</button>`:"0"}</td><td class="num" style="text-align:right">${itemsOn}</td><td>${over?chipFor("Overdue")+` <span class="small num">${over}</span>`:'<span class="muted">—</span>'}</td><td class="num" style="text-align:right">${money(billed)}</td><td><div style="display:flex;gap:6px;justify-content:flex-end">${kids?`<button class="btn sm" data-act="loc-open" data-id="${l.id}">Open</button>`:`<button class="btn sm" data-act="loc-orders" data-id="${l.id}">Orders</button>`}${self?"":`<button class="btn sm" data-act="new-loc" data-parent="${l.id}" title="Add a sub-location under ${esc(l.name)}">+ Sub</button>`}<button class="btn sm" data-act="edit-loc" data-id="${l.id}">Edit</button></div></td></tr>`).join("")}</tbody></table></div></div>`;
}
ui.cust = {tab:"active"};
const custOpenOrders = id => S.orders.filter(o=>o.customerId===id && ACTIVE.includes(o.status));
function vCustomers(){
  const isAdmin = me()?.role==="Admin", tab = ui.cust.tab;
  const live = S.customers.filter(c=>!c.deleted), act = live.filter(c=>!c.archived), arch = live.filter(c=>c.archived);
  const rows = (tab==="archived"?arch:act).map(c=>{ const os = S.orders.filter(o=>o.customerId===c.id);
    return {c,active:os.filter(o=>o.status==="Rented Out").length,billed:os.filter(o=>o.status!=="Cancelled").reduce((s,o)=>s+orderTotals(o).total,0),last:os.map(o=>o.out).sort().pop()}; }).sort((a,b)=>a.c.name.localeCompare(b.c.name));
  return band("Customers","Project managers and superintendents who order equipment.",`<button class="btn primary" data-act="new-cust">+ Add customer</button>`) +
  tabs([["active","Active",act.length],["archived","Archived",arch.length]],tab,"ctab") + `
  <div class="content">${tab==="archived"?`<div class="infobox">Archived customers are hidden from new orders and check-outs. Their past orders, invoices and reports keep their name.${isAdmin?" To use one again, open it with Edit and choose Restore.":""}</div>`:""}
  <div class="tablewrap"><table><thead><tr><th>Name</th><th>Position</th><th>Company</th><th>Phone</th><th>Email</th><th style="text-align:right">Active orders</th><th>Last order</th><th style="text-align:right">Total billed</th><th></th></tr></thead>
  <tbody>${rows.length?rows.map(({c,active,billed,last})=>`<tr><td><b>${esc(c.name)}</b>${c.archived?`<div class="small muted">Archived ${c.archivedAt?fmtDate(c.archivedAt):""}${c.archivedBy?" by "+esc(c.archivedBy):""}</div>`:""}</td><td>${esc(c.position)||'<span class="muted">—</span>'}</td><td>${esc(c.company)||"—"}</td><td class="num">${esc(c.phone)||'<span class="muted">—</span>'}</td><td>${esc(c.email)||'<span class="muted">—</span>'}</td><td class="num" style="text-align:right">${active}</td><td class="num">${last?fmtDate(last):"—"}</td><td class="num" style="text-align:right">${money(billed)}</td>
  <td><div style="display:flex;gap:6px;justify-content:flex-end"><button class="btn sm" data-act="cust-orders" data-id="${c.id}">Orders</button><button class="btn sm" data-act="edit-cust" data-id="${c.id}">Edit</button></div></td></tr>`).join(""):`<tr><td colspan="9"><div class="empty"><strong>${tab==="archived"?"No archived customers":"No customers yet"}</strong>${tab==="archived"?"Customers you archive show up here.":"Add the project managers and superintendents who order equipment."}</div></td></tr>`}</tbody></table></div></div>`;
}
function archiveCustomer(id, on){
  const c = custById(id), open = custOpenOrders(id);
  if (on && open.length){ toast(`${c.name} has ${open.length} open order${open.length>1?"s":""} (#${open.map(o=>o.no).join(", #")}). Return or cancel ${open.length>1?"them":"it"} first.`); return; }
  if (on){ c.archived = true; c.archivedAt = toLocalInput(now()); c.archivedBy = myName(); } else { delete c.archived; delete c.archivedAt; delete c.archivedBy; }
  save(); closeModal(); rerender(); toast(on ? `${c.name} archived` : `${c.name} restored`);
}
function deleteCustomer(id){
  const c = custById(id), open = custOpenOrders(id), past = S.orders.filter(o=>o.customerId===id).length;
  if (open.length){ toast(`${c.name} has ${open.length} open order${open.length>1?"s":""} (#${open.map(o=>o.no).join(", #")}). Return or cancel ${open.length>1?"them":"it"} first.`); return; }
  confirmBox(`Delete ${esc(c.name)}?`, past
    ? `${esc(c.name)} is removed from your customer list and can't be restored. Their ${past} past order${past>1?"s":""} and invoices keep the name. To keep them restorable instead, use <b>Archive</b>.`
    : `${esc(c.name)} has no orders, so they're removed completely. This can't be undone.`, "Delete customer", ()=>{
      if (past){ c.deleted = true; c.deletedAt = toLocalInput(now()); c.deletedBy = myName(); delete c.archived; }
      else S.customers = S.customers.filter(x=>x.id!==id);
      save(); closeModal(); rerender(); toast(`${c.name} deleted`); });
}

/* ---------------- Reports ---------------- */
function periodRange(p){ const d = now(); if (p==="month") return [new Date(d.getFullYear(),d.getMonth(),1), d];
  if (p==="90") return [new Date(d-90*DAY), d]; if (p==="ytd") return [new Date(d.getFullYear(),0,1), d]; return [new Date(2000,0,1), d]; }
/* ================= Reports (each one exports to CSV and Excel) ================= */
ui.reports.tab = "onrent"; ui.reports.loc = ""; ui.reports.withCo = true; ui.reports.lateOnly = false;
/* ---------------- Available report (every role) ---------------- */
ui.av = {tab:"assets", q:"", cat:"", all:false};
function nextBooking(id){ return S.orders.filter(o=>o.status==="Booked" && o.lines.some(l=>l.itemId===id)).sort((a,b)=>parseD(a.out)-parseD(b.out))[0]; }
function availData(){
  const U = ui.av, inGroup = x => !U.cat || x.category===U.cat || groupOf(x.category)?.parent===U.cat;
  const match = (x, extra="") => !U.q || [x.id, x.name, x.category, x.make, x.model, x.sku, keptAt(x), extra].join(" ").toLowerCase().includes(U.q.toLowerCase());
  const fleet = S.items.filter(i=>i.status!=="Retired" && !i.shopOnly);
  const assets = fleet.filter(i=>itemStatus(i)==="Available").sort((a,b)=>az(a.id,b.id));
  const stockAll = S.stock.filter(s=>!s.retired).map(s=>({s, out:stockOut(s.id), avail:Math.max(0,(+s.qty||0)-stockOut(s.id))})).sort((a,b)=>nmSort(a.s,b.s));
  const invAll = S.inventory.filter(i=>!i.retired).sort(nmSort);
  const stock = stockAll.filter(r=>U.all || r.avail>0), inv = invAll.filter(i=>U.all || (+i.qty||0)>0);
  const rates = x => [+x.day||"", +x.week||"", +x.month||""];
  const sheets = [
    {name:"Equipment", cols:[["Asset #","s",9],["Asset","s",30],["Group","s",22],["Make","s",14],["Model","s",14],["Year","s",7],["Kept at","s",26],["Next booking","s",22],["Day","$",10],["Week","$",10],["Month","$",11]],
      rows:assets.filter(i=>inGroup(i)&&match(i)).map(i=>{ const b = nextBooking(i.id); return [i.id, i.name, i.category||"", i.make||"", i.model||"", i.year||"", keptAt(i), b?`#${b.no} from ${fmtDate(b.out)}`:"", ...rates(i)]; })},
    {name:"Asset stock", cols:[["Item","s",30],["Group","s",22],["Kept at","s",26],["Total","n",9],["Checked out","n",11],["Available","n",10],["Unit","s",7],["Day","$",10],["Week","$",10],["Month","$",11]],
      rows:stock.filter(r=>inGroup(r.s)&&match(r.s)).map(r=>[r.s.name, r.s.category||"", keptAt(r.s), +r.s.qty||0, r.out, r.avail, r.s.unit||"ea", ...rates(r.s)])},
    {name:"Inventory", cols:[["Item","s",30],["SKU","s",12],["Group","s",22],["Kept at","s",26],["On hand","n",10],["Unit","s",7],["Status","s",12],["Sold at","$",11]],
      rows:inv.filter(i=>inGroup(i)&&match(i)).map(i=>[i.name, i.sku||"", i.category||"", keptAt(i), +i.qty||0, i.unit||"ea", invStatus(i), canSell(i)?+i.price||0:""])}];
  return {fleet, assets, stockAll, stock, invAll, inv, sheets, srcs:[assets, stock.map(r=>r.s), inv]};
}
function availBtns(){ return `<button class="btn" data-act="av-csv">${ico("download")} CSV</button><button class="btn primary" data-act="av-xlsx">${ico("download")} Excel (all 3)</button>`; }
const AVAIL_SUB = "What's ready to go right now. Equipment that's rented out, checked out, in maintenance or shop use only is left out.";
function availBody(){
  const U = ui.av, D = availData(), idx = {assets:0, stock:1, inv:2}[U.tab] ?? 0, sh = D.sheets[idx];
  const src = D.srcs[idx], cats = [...new Set(src.map(x=>x.category).filter(Boolean))];
  const units = D.stockAll.reduce((s,r)=>s+r.avail,0), low = D.invAll.filter(i=>invStatus(i)!=="In stock").length;
  const cell = (v,i) => { const [h,t] = sh.cols[i];
    const out = v==null||v==="" ? '<span class="muted">—</span>' : t==="$" ? money0(v) : h==="Asset #" ? (isSuper() ? `#${esc(v)}` : `<a href="#asset-${esc(v)}">#${esc(v)}</a>`) : h==="Status" ? `<span class="chip ${v==="In stock"?"available":v==="Low stock"?"maint":"overdue"}">${esc(v)}</span>` : h==="Available" ? `<b>${v}</b>` : esc(String(v));
    return `<td ${t==="$"||t==="n"?'class="num" style="text-align:right"':""}>${out}</td>`; };
  const noun = idx===0 ? "pieces of equipment" : idx===1 ? "stock items" : "inventory items";
  return `
    <div class="kpis">
      <div class="kpi"><span class="lbl">Equipment available</span><span class="val num">${D.assets.length}</span><span class="note">of ${D.fleet.length} in the rental fleet</span></div>
      <div class="kpi"><span class="lbl">Asset stock units</span><span class="val num">${units}</span><span class="note">across ${D.stockAll.filter(r=>r.avail>0).length} items</span></div>
      <div class="kpi"><span class="lbl">Inventory in stock</span><span class="val num">${D.invAll.filter(i=>(+i.qty||0)>0).length}</span><span class="note">of ${D.invAll.length} items</span></div>
      <div class="kpi ${low?"bad":""}"><span class="lbl">Low or out of stock</span><span class="val num">${low}</span><span class="note">inventory to reorder</span></div>
    </div>
    <div class="seg" role="group" aria-label="Show" style="margin-top:14px">${[["assets","Equipment",D.assets.length],["stock","Asset stock",D.stock.length],["inv","Inventory",D.inv.length]].map(([k,l,n])=>`<button class="${U.tab===k?"on":""}" data-act="avtab" data-v="${k}">${l} <span class="num">${n}</span></button>`).join("")}</div>
    <div class="toolbar" style="margin:12px 0;gap:8px;flex-wrap:wrap">
      <div class="search" style="max-width:340px">${ico("search")}<input id="av-q" type="search" placeholder="${idx===0?"Asset #, name, make, model or location":"Search by name or location"}" value="${esc(U.q)}" data-input="avq" aria-label="Search"></div>
      <select data-change="avcat" aria-label="Group"><option value="">All groups</option>${orderedGroups().filter(g=>cats.includes(g.name) || cats.some(c=>groupOf(c)?.parent===g.name)).map(g=>`<option value="${esc(g.name)}" ${U.cat===g.name?"selected":""}>${g.parent?"   › ":""}${esc(g.name)}</option>`).join("")}</select>
      ${idx>0?`<label class="small" style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-change="avall" ${U.all?"checked":""}> Include items with none available</label>`:""}
      <span class="small muted">${sh.rows.length} ${noun}</span>
    </div>
    <div class="tablewrap"><table><thead><tr>${sh.cols.map(([h,t])=>`<th ${t==="$"||t==="n"?'style="text-align:right"':""}>${h}</th>`).join("")}</tr></thead><tbody>${sh.rows.length ? sh.rows.map(r=>`<tr>${r.map(cell).join("")}</tr>`).join("") : `<tr><td colspan="${sh.cols.length}"><div class="empty"><strong>Nothing matches</strong>Try another search or group.</div></td></tr>`}</tbody></table></div>
    <p class="small muted" style="margin-top:10px">As of ${fmtDate(nowStr())} ${fmtTime(nowStr().slice(11,16))}. ${idx===0?"“Next booking” means it's free now but reserved on a booked order starting that day.":idx===1?"Available is the total minus what's checked out.":"Sold at shows the sale price for items that can be sold."}</p>`;
}
const REPORTS = [["onrent","Equipment on rent"],["avail","Available"],["revenue","Revenue"],["util","Utilization"],["cust","Customers"]];
const nowStr = () => toLocalInput(now());
const daysBetween = (a, b) => Math.max(0, Math.floor((parseD(b) - parseD(a)) / DAY));

/* Each builder returns {sheets:[{name, cols:[[header,type,width]], rows:[[...]], totals}] , ...view data}.
   Column types: s text, n number, $ money, d date+time, D date, % percent (0..1). */
function rptOnRent(){
  const u = ui.reports, n = nowStr(), rows = [];
  S.orders.filter(o=>o.status==="Rented Out").forEach(o=>{ const c = custById(o.customerId), l = locById(o.locationId);
    o.lines.forEach(ln=>{ const it = itemById(ln.itemId);
      rows.push({loc:locLabel(l), locId:o.locationId, code:l?.code||"", id:ln.itemId, name:it?.name||"(deleted)", group:it?.category||"", src:"Rental", ref:"#"+o.no, no:o.no, who:c?.name||"", out:o.out, due:o.ret,
        days:daysBetween(o.out,n), late:isOverdue(o), lateDays:isOverdue(o)?daysBetween(o.ret,n):0, month:+ln.month||0, accrued:priceLine(ln,o.out,n).total, booked:priceLine(ln,o.out,o.ret).total, qty:1}); }); });
  if (u.withCo) S.custody.filter(c=>!c.inAt).forEach(c=>{ const l = locById(c.locationId), it = c.kind==="asset" ? itemById(c.itemId) : null, st = c.kind==="stock" ? stockById(c.stockId) : null;
    rows.push({loc:locLabel(l), locId:c.locationId, code:l?.code||"", id:c.kind==="asset"?c.itemId:"", name:it?.name||st?.name||(c.kind==="inv"?S.inventory.find(x=>x.id===c.invId)?.name:"")||"(deleted)", group:it?.category||st?.category||(c.kind==="inv"?"Inventory":""), src:"Check-out", ref:c.ref, who:c.holder, out:c.outAt, due:c.dueAt||"",
      days:daysBetween(c.outAt,n), late:custodyOverdue(c), lateDays:custodyOverdue(c)?daysBetween(c.dueAt,n):0, month:0, accrued:0, booked:0, qty:+c.qty||1, cid:c.id}); });
  let list = rows; if (u.loc) list = list.filter(r=>r.locId===u.loc); if (u.lateOnly) list = list.filter(r=>r.late);
  list.sort((a,b)=>a.loc.localeCompare(b.loc) || (b.late-a.late) || a.due.localeCompare(b.due));
  const byLoc = {}; list.forEach(r=>{ const g = byLoc[r.locId] ||= {loc:r.loc, locId:r.locId, rows:[], rent:0, co:0, late:0, month:0, accrued:0, oldest:r.out}; g.rows.push(r); if (r.src==="Rental") g.rent++; else g.co += r.qty; if (r.late) g.late++; g.month += r.month; g.accrued += r.accrued; if (r.out < g.oldest) g.oldest = r.out; });
  const groups = Object.values(byLoc).sort((a,b)=>b.rows.length-a.rows.length);
  const sheets = [
    {name:"On rent by location", totals:true, cols:[["Location","s",30],["Asset #","s",10],["Description","s",34],["Group","s",20],["Type","s",11],["Order / ref","s",12],["Customer / assigned to","s",22],["Out","d",17],["Due back","d",17],["Days out","n",9],["Status","s",10],["Days late","n",9],["Qty","n",6],["Month rate","$",12],["Rental to date","$",14]],
      rows:list.map(r=>[r.loc,r.id,r.name,r.group,r.src,r.ref,r.who,r.out,r.due,r.days,r.late?"Overdue":r.src==="Rental"?"Rented out":"Checked out",r.lateDays,r.qty,r.src==="Rental"?r.month:null,r.src==="Rental"?r.accrued:null])},
    {name:"Summary by location", totals:true, cols:[["Location","s",30],["Assets on rent","n",14],["Checked out","n",12],["Overdue","n",10],["Month rate value","$",16],["Rental to date","$",15],["Out since","d",17]],
      rows:groups.map(g=>[g.loc,g.rent,g.co,g.late,g.month,g.accrued,g.oldest])}];
  return {sheets, list, groups, all:rows};
}
function rptRevenue(){
  const [a,b] = periodRange(ui.reports.period), d = now();
  const ords = S.orders.filter(o=>o.status!=="Cancelled" && parseD(o.out)>=a && parseD(o.out)<=b);
  const months = []; for (let i=11;i>=0;i--){ const m = new Date(d.getFullYear(), d.getMonth()-i, 1); months.push({m, v:0, r:0, dl:0, c:0, n:0}); }
  S.orders.filter(o=>o.status!=="Cancelled").forEach(o=>{ const od = parseD(o.out); const mm = months.find(x=>x.m.getFullYear()===od.getFullYear()&&x.m.getMonth()===od.getMonth()); if (mm){ const t = orderTotals(o); mm.v += t.total; mm.r += t.rental - t.bundleDisc; mm.dl += t.delivery; mm.c += t.charges + (+o.adjustment||0); mm.n++; } });
  const byLoc = {}; ords.forEach(o=>{ const k = locLabel(locById(o.locationId)); const e = byLoc[k] ||= {n:0,v:0}; e.n++; e.v += orderTotals(o).total; });
  const tot = ords.reduce((s,o)=>{const t=orderTotals(o); s.r+=t.rental-t.bundleDisc; s.d+=t.delivery; s.c+=t.charges+(+o.adjustment||0); return s;},{r:0,d:0,c:0});
  const locRows = Object.entries(byLoc).sort((x,y)=>y[1].v-x[1].v);
  const sheets = [
    {name:"Billed by month", totals:true, cols:[["Month","s",14],["Orders","n",9],["Equipment rental","$",16],["Delivery","$",12],["Other charges","$",14],["Total billed","$",14]],
      rows:months.map(x=>[x.m.toLocaleString("en-US",{month:"short",year:"numeric"}),x.n,x.r,x.dl,x.c,x.v])},
    {name:"Billed by job", totals:true, cols:[["Job","s",32],["Orders","n",9],["Total billed","$",14]], rows:locRows.map(([k,v])=>[k,v.n,v.v])},
    {name:"Orders in period", totals:true, cols:[["Order #","n",9],["Customer","s",20],["Job","s",30],["Items","s",40],["Out","d",17],["Back","d",17],["Status","s",11],["Rental","$",12],["Delivery","$",11],["Charges","$",11],["Discounts & adj.","$",15],["Total","$",13]],
      rows:ords.slice().sort((x,y)=>x.out.localeCompare(y.out)).map(o=>{const t=orderTotals(o); return [o.no,custById(o.customerId)?.name||"",locLabel(locById(o.locationId)),o.lines.map(l=>"#"+l.itemId+" "+(itemById(l.itemId)?.name||"")).join("; "),o.out,orderRentalEnd(o),orderStatus(o),t.rental,t.delivery,t.charges,(+o.adjustment||0)-t.bundleDisc,t.total];})}];
  return {sheets, months, locRows, tot, ords};
}
function rptUtil(){
  const [a] = periodRange(ui.reports.period);
  const rows = S.items.filter(i=>i.status!=="Retired" && !i.shopOnly).map(it=>({it,u:utilization(it.id),cu:custodyUse(it.id),st:itemStatus(it),rev:S.orders.filter(o=>o.status!=="Cancelled"&&parseD(o.out)>=a).reduce((s,o)=>s+orderTotals(o).lines.filter(l=>l.itemId===it.id).reduce((x,l)=>x+l.total,0),0)})).sort((x,y)=>y.u-x.u||y.rev-x.rev);
  const cats = {}; rows.forEach(r=>{const c=cats[r.it.category] ||= {n:0,u:0,rev:0,out:0}; c.n++; c.u+=r.u; c.rev+=r.rev; if (r.st==="Rented Out"||r.st==="Overdue") c.out++;});
  const catRows = Object.entries(cats).sort((x,y)=>y[1].u/y[1].n-x[1].u/x[1].n);
  const sheets = [
    {name:"Utilization by asset", totals:false, cols:[["Asset #","s",10],["Description","s",34],["Group","s",22],["Time on rent (90 days)","%",18],["Crew check-out time (90 days)","%",22],["Rental billed (period)","$",18],["Status","s",14],["Last used","D",12]],
      rows:rows.map(r=>[r.it.id,r.it.name,r.it.category,r.u,r.cu,r.rev,r.st,lastUsed(r.it.id)||""])},
    {name:"Utilization by group", totals:false, cols:[["Group","s",24],["Assets","n",9],["On rent now","n",12],["Avg time on rent (90 days)","%",22],["Rental billed (period)","$",18]],
      rows:catRows.map(([k,v])=>[k,v.n,v.out,v.u/v.n,v.rev])}];
  return {sheets, rows, catRows};
}
function rptCust(){
  const [a,b] = periodRange(ui.reports.period);
  const rows = S.customers.filter(c=>!c.deleted || S.orders.some(o=>o.customerId===c.id && o.status!=="Cancelled")).map(c=>{ const all = S.orders.filter(o=>o.customerId===c.id && o.status!=="Cancelled"), os = all.filter(o=>parseD(o.out)>=a && parseD(o.out)<=b);
    return {c, n:os.length, v:os.reduce((s,o)=>s+orderTotals(o).total,0), active:all.filter(o=>o.status==="Rented Out").length, late:all.filter(isOverdue).length, last:all.map(o=>o.out).sort().pop()||"", life:all.reduce((s,o)=>s+orderTotals(o).total,0)}; }).sort((x,y)=>y.v-x.v||y.life-x.life);
  const sheets = [{name:"Billed by customer", totals:true, cols:[["Customer","s",22],["Position","s",24],["Company","s",18],["Phone","s",14],["Email","s",26],["Orders (period)","n",14],["Billed (period)","$",15],["On rent now","n",12],["Overdue now","n",12],["Last order","d",17],["Billed all time","$",15]],
    rows:rows.map(r=>[r.c.name,r.c.position||"",r.c.company,r.c.phone,r.c.email,r.n,r.v,r.active,r.late,r.last,r.life])}];
  return {sheets, rows};
}
const RPT = {onrent:rptOnRent, revenue:rptRevenue, util:rptUtil, cust:rptCust};

function sheetTable(sh, max){
  const fmt = (v,t) => v==null||v==="" ? '<span class="muted">—</span>' : t==="$" ? money(v) : t==="%" ? Math.round(v*100)+"%" : t==="d" ? fmtDT(v).replace("<br>"," ") : t==="D" ? fmtDate(v) : esc(v);
  const rows = max ? sh.rows.slice(0,max) : sh.rows;
  const right = t => "$n%".includes(t) ? ' style="text-align:right"' : "";
  return `<div class="tablewrap" style="border:0"><table><thead><tr>${sh.cols.map(([h,t])=>`<th${right(t)}>${esc(h)}</th>`).join("")}</tr></thead><tbody>${rows.length?rows.map(r=>`<tr>${r.map((v,i)=>`<td class="${"$n%dD".includes(sh.cols[i][1])?"num":""}"${right(sh.cols[i][1])}>${fmt(v,sh.cols[i][1])}</td>`).join("")}</tr>`).join(""):`<tr><td colspan="${sh.cols.length}"><div class="empty">Nothing in this report for the period.</div></td></tr>`}
  ${sh.totals&&rows.length?`<tr class="tot">${sh.cols.map(([h,t],i)=>`<td${right(t)} class="num">${i===0?"<b>Total</b>":(t==="$"||t==="n")&&!/^(Days|Order #)/.test(h)?`<b>${t==="$"?money(sh.rows.reduce((s,r)=>s+(+r[i]||0),0)):sh.rows.reduce((s,r)=>s+(+r[i]||0),0)}</b>`:""}</td>`).join("")}</tr>`:""}</tbody></table></div>`;
}
function vReports(){
  const u = ui.reports;
  if (!can("reports")) u.tab = "avail";
  if (u.tab==="avail") return band("Reports", AVAIL_SUB, availBtns()) + tabs(can("reports") ? REPORTS : [["avail","Available"]], u.tab, "rtab") + `<div class="content">${availBody()}</div>`;
  const R = RPT[u.tab](), periodOn = u.tab!=="onrent";
  const periodSel = periodOn ? `<select id="rperiod" data-change="rperiod" class="btn" aria-label="Period">${[["month","This month"],["90","Last 90 days"],["ytd","Year to date"],["all","All time"]].map(([k,l])=>`<option value="${k}" ${u.period===k?"selected":""}>${l}</option>`).join("")}</select>` : "";
  const exp = `${periodSel}<button class="btn" data-act="rpt-csv" data-i="0">${ico("download")} CSV</button><button class="btn" data-act="rpt-xlsx">${ico("download")} Excel</button><button class="btn primary" data-act="rpt-all">All reports to Excel</button>`;
  const csvBtn = i => `<button class="btn sm" data-act="rpt-csv" data-i="${i}">${ico("download")} CSV</button>`;
  let body = "";
  if (u.tab==="onrent"){
    const all = R.all, rentals = all.filter(r=>r.src==="Rental");
    const locs = [...new Map(all.map(r=>[r.locId,r.loc])).entries()].sort((a,b)=>a[1].localeCompare(b[1]));
    const maxN = Math.max(1,...R.groups.map(g=>g.rows.length));
    body = `<div class="kpis">
      <div class="kpi"><span class="lbl">Assets on rent</span><span class="val num">${rentals.length}</span><span class="note">across ${new Set(rentals.map(r=>r.locId)).size} job${new Set(rentals.map(r=>r.locId)).size===1?"":"s"}</span></div>
      <div class="kpi ${rentals.some(r=>r.late)?"bad":""}"><span class="lbl">Overdue</span><span class="val num">${rentals.filter(r=>r.late).length}</span><span class="note">past the booked return</span></div>
      <div class="kpi"><span class="lbl">Month rate value out</span><span class="val num">${money0(rentals.reduce((s,r)=>s+r.month,0))}</span><span class="note">at each asset's month rate</span></div>
      <div class="kpi"><span class="lbl">Rental accrued to date</span><span class="val num">${money0(rentals.reduce((s,r)=>s+r.accrued,0))}</span><span class="note">rent-out to today, cheapest rate mix</span></div>
      ${u.withCo?`<div class="kpi"><span class="lbl">Crew check-outs</span><span class="val num">${all.filter(r=>r.src==="Check-out").length}</span><span class="note">not billed</span></div>`:""}
    </div>
    <div class="toolbar">
      <select id="rloc" data-change="rloc" aria-label="Location"><option value="">All locations</option>${locs.map(([id,l])=>`<option value="${id}" ${u.loc===id?"selected":""}>${esc(l)}</option>`).join("")}</select>
      <label class="check"><input type="checkbox" id="rco" data-change="rco" ${u.withCo?"checked":""}> Include crew check-outs</label>
      <label class="check"><input type="checkbox" id="rlate" data-change="rlate" ${u.lateOnly?"checked":""}> Overdue only</label>
    </div>
    <div class="panel"><h3>Summary by location ${csvBtn(1)}</h3><div class="body"><div class="bars">${R.groups.length?R.groups.map(g=>`<div class="bar-row loc-bar"><button class="linkbtn" data-act="rloc-pick" data-id="${g.locId}">${esc(g.loc)}</button><div class="track"><div class="fill" style="width:${g.rows.length/maxN*100}%"></div></div><span class="v">${g.rent?g.rent+" on rent":""}${g.rent&&g.co?" · ":""}${g.co?g.co+" checked out":""}${g.late?` · <b style="color:var(--bad)">${g.late} late</b>`:""}</span></div>`).join(""):'<div class="muted">Nothing is out right now.</div>'}</div></div></div>
    ${R.groups.map(g=>`<section class="panel"><h3><span>${esc(g.loc)} <span class="muted small" style="font-weight:400">${g.rows.length} item${g.rows.length===1?"":"s"}${g.month?` · ${money0(g.month)}/mo`:""} · out since ${fmtDate(g.oldest)}</span></span><button class="btn sm" data-act="loc-orders" data-id="${g.locId}">Orders</button></h3>
      <div class="tablewrap" style="border:0"><table class="loc-table"><colgroup><col style="width:22%"><col style="width:12%"><col style="width:7%"><col style="width:12%"><col style="width:8%"><col style="width:8%"><col style="width:6%"><col style="width:9%"><col style="width:8%"><col style="width:8%"></colgroup><thead><tr><th>Asset</th><th>Group</th><th>Order / ref</th><th>Customer / assigned to</th><th>Out</th><th>Due back</th><th style="text-align:right">Days out</th><th>Status</th><th style="text-align:right">Month rate</th><th style="text-align:right">Rental to date</th></tr></thead><tbody>
      ${g.rows.map(r=>`<tr><td>${r.id?`<button class="linkbtn" data-act="view-item" data-id="${esc(r.id)}">#${esc(r.id)}</button> `:""}${esc(r.name)}${r.qty>1?` × ${r.qty}`:""}</td><td class="muted">${esc(r.group)}</td><td class="num">${r.no?`<button class="linkbtn" data-act="view-order" data-no="${r.no}">${esc(r.ref)}</button>`:esc(r.ref)}</td><td>${esc(r.who)}</td><td class="num">${fmtDate(r.out)}</td><td class="num">${r.due?fmtDate(r.due):'<span class="muted">Open</span>'}</td><td class="num" style="text-align:right">${r.days}</td><td>${r.late?chipFor("Overdue")+` <span class="small num" style="color:var(--bad)">${r.lateDays}d</span>`:chipFor(r.src==="Rental"?"Rented Out":"Checked Out")}</td><td class="num" style="text-align:right">${r.src==="Rental"?money(r.month):'<span class="muted">—</span>'}</td><td class="num" style="text-align:right">${r.src==="Rental"?money(r.accrued):'<span class="muted">not billed</span>'}</td></tr>`).join("")}
      </tbody></table></div></section>`).join("")}`;
  } else if (u.tab==="revenue"){
    const months = R.months, nice = niceMax(Math.max(1,...months.map(x=>x.v)));
    const W=720,H=220,L=56,B=28,T=12,bw=(W-L-10)/12;
    const svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Billed by month">${[0,.25,.5,.75,1].map(f=>{const t=f*nice, y=H-B-f*(H-B-T); return `<line x1="${L}" x2="${W-6}" y1="${y}" y2="${y}" stroke="var(--line)" /><text x="${L-6}" y="${y+3}" text-anchor="end">${money0(t)}</text>`;}).join("")}
      ${months.map((x,i)=>{const h=(x.v/nice)*(H-B-T), X=L+i*bw+bw*.18; return `<rect x="${X}" y="${H-B-h}" width="${bw*.64}" height="${Math.max(0,h)}" fill="var(--accent)" rx="2"><title>${x.m.toLocaleString("en-US",{month:"short",year:"numeric"})}: ${money(x.v)}</title></rect><text x="${X+bw*.32}" y="${H-B+14}" text-anchor="middle">${x.m.toLocaleString("en-US",{month:"short"})}</text>${x.v?`<text x="${X+bw*.32}" y="${H-B-h-4}" text-anchor="middle" style="fill:var(--fg)">${money0(x.v)}</text>`:""}`;}).join("")}</svg>`;
    body = `<div class="kpis"><div class="kpi"><span class="lbl">Total billed</span><span class="val num">${money0(R.tot.r+R.tot.d+R.tot.c)}</span><span class="note">${R.ords.length} orders by rent-out date</span></div><div class="kpi"><span class="lbl">Equipment rental</span><span class="val num">${money0(R.tot.r)}</span></div><div class="kpi"><span class="lbl">Delivery</span><span class="val num">${money0(R.tot.d)}</span></div><div class="kpi"><span class="lbl">Other charges</span><span class="val num">${money0(R.tot.c)}</span><span class="note">fuel, cleaning, mechanic, adjustments</span></div></div>
    <div class="panel"><h3>Billed by month (last 12 months) ${csvBtn(0)}</h3><div class="body chart">${svg}</div></div>
    <div class="panel"><h3>Billed by job ${csvBtn(1)}</h3>${sheetTable(R.sheets[1])}</div>
    <div class="panel"><h3>Orders in period ${csvBtn(2)}</h3>${sheetTable(R.sheets[2])}</div>`;
  } else if (u.tab==="util"){
    body = `<div class="panel"><h3>Time on rent by group, last 90 days ${csvBtn(1)}</h3><div class="body"><div class="bars">${R.catRows.map(([k,v])=>`<div class="bar-row"><span>${esc(k)}</span><div class="track"><div class="fill" style="width:${v.u/v.n*100}%"></div></div><span class="v">${Math.round(v.u/v.n*100)}%</span></div>`).join("")}</div></div></div>
    <div class="panel"><h3>By asset ${csvBtn(0)}</h3>${sheetTable(R.sheets[0])}</div>`;
  } else {
    const max = Math.max(1,...R.rows.map(r=>r.v));
    body = `<div class="panel"><h3>Billed by customer ${csvBtn(0)}</h3><div class="body"><div class="bars">${R.rows.filter(r=>r.n).length?R.rows.filter(r=>r.n).map(r=>`<div class="bar-row"><span>${esc(r.c.name)} <span class="muted small">(${r.n})</span></span><div class="track"><div class="fill" style="width:${r.v/max*100}%"></div></div><span class="v">${money0(r.v)}</span></div>`).join(""):'<div class="muted">No orders in this period.</div>'}</div></div>${sheetTable(R.sheets[0])}</div>`;
  }
  return band("Reports", u.tab==="onrent"?`Where every piece of equipment is right now, as of ${fmtDate(nowStr())} ${nowStr().slice(11)}.`:"", exp) + tabs(REPORTS,u.tab,"rtab") + `<div class="content">${body}</div>`;
}

/* ---- Exports ---- */
const stamp = () => nowStr().slice(0,10);
const slug = s => s.toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"");
function sheetCsv(sh){
  const f = (v,t) => v==null||v==="" ? "" : t==="d" ? v.replace("T"," ") : t==="D" ? v.slice(0,10) : t==="%" ? Math.round(v*1000)/10+"%" : t==="$" ? (+v).toFixed(2) : v;
  const rows = [sh.cols.map(c=>c[0])].concat(sh.rows.map(r=>r.map((v,i)=>f(v,sh.cols[i][1]))));
  if (sh.totals && sh.rows.length) rows.push(sh.cols.map(([h,t],i)=>i===0?"Total":(t==="$"||t==="n")&&!/^(Days|Order #)/.test(h)?(t==="$"?sh.rows.reduce((s,r)=>s+(+r[i]||0),0).toFixed(2):sh.rows.reduce((s,r)=>s+(+r[i]||0),0)):""));
  return csv(rows);
}
function exportReportCsv(i){ const tab = ui.reports.tab, R = RPT[tab](), sh = R.sheets[+i] || R.sheets[0]; download(`BH-${slug(sh.name)}-${stamp()}.csv`, sheetCsv(sh)); }
function reportHeader(){ const u = ui.reports; return u.tab==="onrent" ? `As of ${fmtDate(nowStr())} ${nowStr().slice(11)}${u.loc?` · ${locLabel(locById(u.loc))}`:""}${u.lateOnly?" · overdue only":""}` : `Period: ${{month:"This month",90:"Last 90 days",ytd:"Year to date",all:"All time"}[u.period]} · exported ${fmtDate(nowStr())}`; }
function exportReportXlsx(){ const name = REPORTS.find(r=>r[0]===ui.reports.tab)[1]; saveFile(`BH-${slug(name)}-${stamp()}.xlsx`, buildXlsx(RPT[ui.reports.tab]().sheets)); }
function exportAllXlsx(){ const keep = ui.reports.tab; const sheets = []; REPORTS.forEach(([k])=>{ ui.reports.tab = k; sheets.push(...RPT[k]().sheets); }); ui.reports.tab = keep; saveFile(`BH-all-reports-${stamp()}.xlsx`, buildXlsx(sheets)); }

/* ---- Minimal .xlsx writer (no library: works offline and inside SharePoint) ---- */
const CRC_T = (()=>{ const t = new Uint32Array(256); for (let n=0;n<256;n++){ let c=n; for (let k=0;k<8;k++) c = c&1 ? 0xEDB88320^(c>>>1) : c>>>1; t[n]=c>>>0; } return t; })();
function crc32(b){ let c = 0xFFFFFFFF; for (let i=0;i<b.length;i++) c = CRC_T[(c^b[i])&255]^(c>>>8); return (c^0xFFFFFFFF)>>>0; }
function zipStore(files){
  const enc = new TextEncoder(), parts = [], central = []; let off = 0;
  const d = new Date(), dosT = (d.getHours()<<11)|(d.getMinutes()<<5)|(d.getSeconds()>>1), dosD = ((d.getFullYear()-1980)<<9)|((d.getMonth()+1)<<5)|d.getDate();
  for (const [name, text] of files){
    const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30)); h.setUint32(0,0x04034b50,true); h.setUint16(4,20,true); h.setUint16(6,0x0800,true); h.setUint16(8,0,true); h.setUint16(10,dosT,true); h.setUint16(12,dosD,true); h.setUint32(14,crc,true); h.setUint32(18,data.length,true); h.setUint32(22,data.length,true); h.setUint16(26,nb.length,true); h.setUint16(28,0,true);
    parts.push(new Uint8Array(h.buffer), nb, data);
    const c = new DataView(new ArrayBuffer(46)); c.setUint32(0,0x02014b50,true); c.setUint16(4,20,true); c.setUint16(6,20,true); c.setUint16(8,0x0800,true); c.setUint16(10,0,true); c.setUint16(12,dosT,true); c.setUint16(14,dosD,true); c.setUint32(16,crc,true); c.setUint32(20,data.length,true); c.setUint32(24,data.length,true); c.setUint16(28,nb.length,true); c.setUint32(42,off,true);
    central.push(new Uint8Array(c.buffer), nb); off += 30 + nb.length + data.length;
  }
  const cdSize = central.reduce((s,p)=>s+p.length,0), e = new DataView(new ArrayBuffer(22));
  e.setUint32(0,0x06054b50,true); e.setUint16(8,files.length,true); e.setUint16(10,files.length,true); e.setUint32(12,cdSize,true); e.setUint32(16,off,true);
  return new Blob([...parts, ...central, new Uint8Array(e.buffer)], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
function buildXlsx(sheets){
  const x = s => String(s).replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,"");
  const col = i => { let s=""; i++; while (i){ const m=(i-1)%26; s=String.fromCharCode(65+m)+s; i=Math.floor((i-1)/26); } return s; };
  const serial = v => { const d = parseD(v.length===10?v+"T00:00":v); return (Date.UTC(d.getFullYear(),d.getMonth(),d.getDate(),d.getHours(),d.getMinutes()) - Date.UTC(1899,11,30)) / DAY; };
  const used = new Set();
  const names = sheets.map(sh=>{ let n = sh.name.replace(/[\[\]:*?\/\\]/g,"").slice(0,31) || "Sheet"; let k=2; while (used.has(n.toLowerCase())) n = n.slice(0,28)+" "+(k++); used.add(n.toLowerCase()); return n; });
  const cell = (ref, v, t, bold) => {
    if (v==null || v==="") return "";
    if (t==="$") return `<c r="${ref}" s="${bold?6:2}"><v>${+v}</v></c>`;
    if (t==="n") return `<c r="${ref}"${bold?' s="8"':""}><v>${+v}</v></c>`;
    if (t==="%") return `<c r="${ref}" s="4"><v>${+v}</v></c>`;
    if (t==="d" || t==="D") return `<c r="${ref}" s="${t==="d"?3:5}"><v>${serial(v)}</v></c>`;
    return `<c r="${ref}" t="inlineStr"${bold?' s="7"':""}><is><t xml:space="preserve">${x(v)}</t></is></c>`;
  };
  const sheetXml = sh => {
    const rows = []; rows.push(`<row r="1">${sh.cols.map((c,i)=>`<c r="${col(i)}1" t="inlineStr" s="1"><is><t>${x(c[0])}</t></is></c>`).join("")}</row>`);
    sh.rows.forEach((r,ri)=>rows.push(`<row r="${ri+2}">${r.map((v,i)=>cell(col(i)+(ri+2), v, sh.cols[i][1])).join("")}</row>`));
    if (sh.totals && sh.rows.length){ const n = sh.rows.length+2;
      rows.push(`<row r="${n}">${sh.cols.map(([h,t],i)=> i===0 ? cell("A"+n,"Total","s",true) : (t==="$"||t==="n")&&!/^(Days|Order #)/.test(h) ? `<c r="${col(i)}${n}" s="${t==="$"?6:8}"><f>SUM(${col(i)}2:${col(i)}${n-1})</f><v>${sh.rows.reduce((s,r)=>s+(+r[i]||0),0)}</v></c>` : "").join("")}</row>`); }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${sh.cols.map((c,i)=>`<col min="${i+1}" max="${i+1}" width="${c[2]||14}" customWidth="1"/>`).join("")}</cols><sheetData>${rows.join("")}</sheetData></worksheet>`;
  };
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`],
    ["docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${x(S.settings.company)} report</dc:title><dc:creator>${x(myName()||S.settings.company)}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0,19)}Z</dcterms:created></cp:coreProperties>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n,i)=>`<sheet name="${x(n)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/><numFmt numFmtId="165" formatCode="mm/dd/yyyy hh:mm"/><numFmt numFmtId="166" formatCode="mm/dd/yyyy"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4B5320"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`],
    ...sheets.map((sh,i)=>[`xl/worksheets/sheet${i+1}.xml`, sheetXml(sh)]),
  ];
  return zipStore(files);
}

function niceMax(v){ const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1,1.5,2,2.5,3,4,5,6,8,10]) if (m*p >= v) return m*p; return 10*p; }

/* ---------------- Setup ---------------- */
function vSetup(){
  const s = S.settings;
  return band("Setup","Company details, delivery and charge rates, and where your data is saved.") + `
  <div class="content">
    <div class="grid2">
      <div class="panel"><h3>Company</h3><div class="body"><div class="form">
        <div class="field full"><label>Company logo</label><div class="photo-row"><div class="logo-prev">${s.logo?`<img src="${s.logo}" alt="Current logo">`:'<span class="small muted">No logo</span>'}</div><div style="display:flex;flex-direction:column;gap:6px"><div style="display:flex;gap:6px;flex-wrap:wrap"><button type="button" class="btn sm" data-act="logo-up">${s.logo?"Change logo":"Upload logo"}</button>${s.logo?`<button type="button" class="btn sm" data-act="logo-del">Remove</button>`:""}</div><span class="hint">PNG, JPG or SVG. A transparent PNG looks best. It appears in the top bar, on the sign-in screen and on invoices.</span></div></div></div>
        <div class="field"><label for="s-company">Company</label><input id="s-company" data-set="company" value="${esc(s.company)}"></div>
        <div class="field"><label for="s-contact">Equipment contact</label><input id="s-contact" data-set="contact" value="${esc(s.contact)}"></div>
        <div class="field"><label for="s-phone">Phone</label><input id="s-phone" data-set="phone" value="${esc(s.phone)}"></div>
        <div class="field"><label for="s-email">Email</label><input id="s-email" data-set="email" value="${esc(s.email)}"></div>
        <div class="field"><label for="s-shopdays">Shop target (days)</label><input id="s-shopdays" type="number" min="1" data-set="shopTargetDays" value="${s.shopTargetDays||30}"><span class="hint">Service requests in the shop longer than this are flagged</span></div>
        <div class="field"><label for="s-cycle">Billing month (days)</label><select id="s-cycle" data-set="cycle"><option value="28" ${s.cycle==28?"selected":""}>28 days (4 weeks)</option><option value="30" ${s.cycle==30?"selected":""}>30 days</option><option value="31" ${s.cycle==31?"selected":""}>31 days</option></select><span class="hint">Orders are billed at the cheapest mix of month, week and day rates.</span></div>
        ${me()?.role==="Admin"?`<div class="field"><label for="s-next">Next order #</label><input id="s-next" type="number" min="1" step="1" data-nextorder="1" value="${S.nextOrder}"><span class="hint">New orders count up from here. Numbers already used are skipped.</span></div>`:""}
        <div class="field full subhead" style="margin:6px 0 -2px"><span>Invoices</span></div>
        <div class="field full"><label for="s-address">Company address (top of invoices)</label><textarea id="s-address" data-set="address" placeholder="Street&#10;City, State ZIP">${esc(s.address||"")}</textarea></div>
        <div class="field"><label for="s-terms">Payment terms</label><select id="s-terms" data-set="terms">${["Due on receipt","Net 15","Net 30","Net 45","Net 60"].map(x=>`<option ${(s.terms||"Net 30")===x?"selected":""}>${x}</option>`).join("")}</select></div>
        <div class="field"><label for="s-tax">Sales tax %</label><input id="s-tax" type="number" step="0.01" min="0" data-set="taxRate" value="${s.taxRate??0}"><span class="hint">Assets with their own tax rate use that instead</span></div>
        <div class="field full"><label for="s-invnote">Invoice note</label><textarea id="s-invnote" data-set="invoiceNote" placeholder="Remit to, job-cost codes, thank-you note">${esc(s.invoiceNote||"")}</textarea></div>
        <div class="field full"><label for="s-policy">Rental policy (for your records, not printed on invoices)</label><textarea id="s-policy" data-set="policy">${esc(s.policy)}</textarea></div>
      </div></div></div>
      <div class="panel"><h3>Your data</h3><div class="body" style="display:flex;flex-direction:column;gap:12px">
        <div class="infobox">Everything saves to the Equipment site in SharePoint as you work. There is no data file to keep.</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn" data-act="export-orders">Orders CSV</button><button class="btn" data-act="export-items">Assets CSV</button></div>
        <div class="small muted">Data size ${fmtBytes(JSON.stringify(S).length)} · ${S.items.length} assets · ${S.orders.length} orders · ${S.customers.length} customers · ${S.locations.length} locations</div>
      </div></div>
    </div>
    <div class="grid2">
      <div class="panel rate-table"><h3>Delivery rates (per trip) <button class="btn sm" data-act="add-rate" data-v="delivery">+ Add</button></h3><div class="tablewrap" style="border:0"><table><thead><tr><th>Zone</th><th>Rate</th><th></th></tr></thead><tbody>${s.delivery.map((d,i)=>`<tr><td><input id="dz-${i}" data-rate="delivery" data-i="${i}" data-f="name" value="${esc(d.name)}" aria-label="Zone"></td><td><input id="dr-${i}" type="number" step="0.01" data-rate="delivery" data-i="${i}" data-f="rate" value="${d.rate}" aria-label="Rate"></td><td><button class="del" data-act="del-rate" data-v="delivery" data-i="${i}" aria-label="Remove">×</button></td></tr>`).join("")}</tbody></table></div></div>
      <div class="panel rate-table"><h3>Other charges <button class="btn sm" data-act="add-rate" data-v="charges">+ Add</button></h3><div class="tablewrap" style="border:0"><table><thead><tr><th>Charge</th><th>Rate</th><th>Per</th><th></th></tr></thead><tbody>${s.charges.map((d,i)=>`<tr><td><input id="cn-${i}" data-rate="charges" data-i="${i}" data-f="name" value="${esc(d.name)}" aria-label="Charge"></td><td><input id="cr-${i}" type="number" step="0.01" data-rate="charges" data-i="${i}" data-f="rate" value="${d.rate}" aria-label="Rate"></td><td><input id="cu-${i}" data-rate="charges" data-i="${i}" data-f="unit" value="${esc(d.unit)}" aria-label="Unit"></td><td><button class="del" data-act="del-rate" data-v="charges" data-i="${i}" aria-label="Remove">×</button></td></tr>`).join("")}</tbody></table></div></div>
    </div>
    ${can("users")?`<div class="panel"><h3>Users &amp; roles <button class="btn sm" data-act="goto" data-to="users">Manage users</button></h3><div class="body small">Managed by Microsoft 365. Everyone signs in with their Bailey-Harris account; open Manage users to see who has access to what.</div></div>`:""}
  </div>`;
}

/* ---------------- Modals ---------------- */
function openModal(html, cls=""){ $("#modal-root").innerHTML = `<div class="scrim" data-act="scrim"><div class="modal ${cls}" role="dialog" aria-modal="true">${html}</div></div>`; applyPerms($("#modal-root")); cardify($("#modal-root")); const f = $("#modal-root input, #modal-root select"); if (f) f.focus(); }
function closeModal(){ $("#modal-root").innerHTML = ""; }
/* a modal with fields to fill in only closes with Cancel, ×, Save or a double Esc, never from a stray click outside it */
const modalHasForm = () => !!$("#modal-root .modal")?.querySelector('input:not([type=search]):not([type=hidden]):not([type=file]), textarea, select');
function nudgeModal(msg){ const m = $("#modal-root .modal"); if (!m) return; m.classList.remove("nudge"); void m.offsetWidth; m.classList.add("nudge"); toast(msg); }
let escArmed = 0;
function toast(msg){ document.querySelectorAll(".toast").forEach(x=>x.remove()); const t = document.createElement("div"); t.className="toast"; t.textContent=msg; document.body.appendChild(t); setTimeout(()=>t.remove(),2600); }
let confirmCb = null;
function confirmBox(title, text, okLabel, cb, danger=true){ confirmCb = cb;
  openModal(`<header><h2>${esc(title)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header><div class="mbody"><p style="margin:0">${text}</p></div><footer><button class="btn" data-act="close">Keep it</button><button class="btn ${danger?"danger":"primary"}" data-act="confirm-ok">${esc(okLabel)}</button></footer>`,"narrow"); }

/* ---- Order form ---- */
let draft = null;
function newOrderDraft(preItem){
  const d = now(); d.setMinutes(0,0,0); d.setHours(d.getHours()+1);
  const r = new Date(d.getTime() + (+S.settings.cycle||28)*DAY);
  return {no:null,customerId:byName(S.customers.filter(c=>!c.archived&&!c.deleted))[0]?.id||"",locationId:sortedLocs().find(l=>!l.isYard)?.id||yardId(),lines:preItem?[lineFrom(preItem)]:[],out:toLocalInput(d),ret:toLocalInput(r),status:"Rented Out",description:"",pickup:"",delivery:{name:"",rate:0,trips:2},charges:[],adjustment:0,notes:"",pq:""};
}
function lineFrom(id){ const it = itemById(id); return {itemId:id,day:it.day,week:it.week,month:it.month}; }
function openOrderForm(o){
  draft = o ? JSON.parse(JSON.stringify(o)) : newOrderDraft();
  draft.pq = draft.pq || "";
  drawOrderForm();
}
function drawOrderForm(){
  const o = draft, t = orderTotals(o), dateOk = parseD(o.ret) > parseD(o.out), valid = dateOk && !!custById(o.customerId);
  const q = (o.pq||"").toLowerCase();
  const pickList = S.items.filter(it=>it.status!=="Retired" && (!it.shopOnly || o.lines.some(l=>l.itemId===it.id))).filter(it=>!q || (it.id+" "+it.name+" "+it.category).toLowerCase().includes(q));
  const pickHtml = pickList.map(it=>{ const sel = o.lines.some(l=>l.itemId===it.id);
    const c = dateOk ? conflicts(it.id,o.out,o.ret,o.no) : []; const m = openMaint(it.id); const cu = openCustody(it.id);
    const busy = !sel && (c.length || m || cu);
    const note = m ? "In maintenance" : cu ? `Checked out to ${cu.holder}` : c.length ? `Booked on #${c[0].no}` : (it.month!=null?money0(it.month)+"/mo":"");
    return `<label class="pick ${busy?"busy":""}"><input type="checkbox" data-act="pick" data-id="${esc(it.id)}" ${sel?"checked":""} ${busy?"disabled":""}><span class="num muted">#${esc(it.id)}</span><span class="trunc" style="max-width:none">${esc(it.name)}</span><span class="small ${busy?"":"muted"} num">${note}</span></label>`; }).join("");
  const days = dateOk ? rentalDays(o.out,o.ret) : 0;
  openModal(`<header><h2>${o.no?`Edit order #${o.no}`:"New order"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">
    <div class="form">
      ${me()?.role==="Admin"?`<div class="field"><label for="f-no">Order #</label><input id="f-no" type="number" min="1" step="1" data-of="newNo" value="${esc(o.newNo ?? (o.no || nextFreeOrderNo()))}"><span class="hint">${o.no?"Admins only. Can't match another order.":"Admins can set a different number."}</span></div>`:""}
      <div class="field"><label for="f-cust">Customer</label><select id="f-cust" data-of="customerId">${custById(o.customerId)?"":`<option value="" selected disabled>Pick a customer…</option>`}${byName(S.customers.filter(c=>(!c.archived&&!c.deleted)||c.id===o.customerId)).map(c=>`<option value="${c.id}" ${o.customerId===c.id?"selected":""}>${esc(c.name)}</option>`).join("")}</select></div>
      <div class="field"><label for="f-loc">Location / job</label><select id="f-loc" data-of="locationId">${sortedLocs().map(l=>`<option value="${l.id}" ${o.locationId===l.id?"selected":""}>${esc(locOpt(l))}</option>`).join("")}</select></div>
      ${!o.no || o.status==="Booked" ? `<div class="field full"><label>When is it going out?</label><div class="seg big" role="radiogroup" aria-label="Book or rent out now">
        <label class="${o.status==="Booked"?"on":""}"><input type="radio" name="f-when" data-of="status" value="Booked" ${o.status==="Booked"?"checked":""}><b>Book it</b><small>Reserve the equipment for these dates. Rent it out when it leaves the yard.</small></label>
        <label class="${o.status==="Rented Out"?"on":""}"><input type="radio" name="f-when" data-of="status" value="Rented Out" ${o.status==="Rented Out"?"checked":""}><b>Rent out now</b><small>The equipment is leaving the yard today.</small></label></div>
        ${o.no?`<span class="hint"><button class="linkbtn" type="button" data-act="cancel-order" data-no="${o.no}">Cancel this booking</button></span>`:""}</div>` : `<div class="field"><label for="f-status">Status</label><select id="f-status" data-of="status">${["Booked","Rented Out"].concat(o.no?["Returned","Cancelled"]:[]).map(s=>`<option ${o.status===s?"selected":""}>${s}</option>`).join("")}</select><span class="hint">${o.status==="Booked"?"Holds the equipment for these dates. Use Rent out now when it leaves the yard.":"The equipment is leaving now."}</span></div>`}
      <div class="field"><label for="f-out">Rent out</label><input id="f-out" type="datetime-local" data-of="out" value="${esc(o.out)}"></div>
      <div class="field"><label for="f-ret">Return date</label><input id="f-ret" type="datetime-local" data-of="ret" value="${esc(o.ret)}"><span class="hint num">${dateOk?`${days} day${days>1?"s":""} billed`:""}</span></div>
      <div class="field"><label for="f-pickup">Pickup status</label><select id="f-pickup" data-of="pickup"><option value="" ${PICKUPS.includes(o.pickup)?"":"selected"}>Not set</option>${PICKUPS.map(s=>`<option ${o.pickup===s?"selected":""}>${s}</option>`).join("")}</select></div>
      <div class="field"><label for="f-cycle">Billing</label><select id="f-cycle" data-of="cycle"><option value="" ${o.cycle?"":"selected"}>One invoice when it comes back</option><option value="cycle" ${o.cycle==="cycle"?"selected":""}>Cycle billing, every ${cycleDays()} days</option><option value="month" ${o.cycle==="month"?"selected":""}>Cycle billing, monthly (same date)</option></select><span class="hint">${o.cycle?"Bill each cycle while the equipment is still on the job. Return charges go on the final invoice.":"Pick cycle billing for long rentals that need a bill every month."}</span></div>
      ${o.bills?.length?`<div class="field full"><div class="warnbox" style="margin:0">${o.bills.length} cycle invoice${o.bills.length>1?"s have":" has"} already gone out. Changing the rent-out date or billing type changes what's left to bill.</div></div>`:""}
      <div class="field full"><label for="f-desc">Description</label><input id="f-desc" data-of="description" value="${esc(o.description)}" placeholder="Shown in the orders list"></div>
    </div>
    ${!custById(o.customerId)?`<div class="warnbox">Pick the customer for this order.</div>`:parseD(o.ret) > parseD(o.out)?"":`<div class="warnbox">The return date has to be after the rent-out date.</div>`}
    <div class="subhead"><span>Equipment</span><span class="small" style="text-transform:none;letter-spacing:0">Booked and out-of-service assets are greyed out for these dates</span></div>
    <div class="field"><input id="f-pq" type="search" data-of="pq" placeholder="Find an asset by number or name" value="${esc(o.pq)}" aria-label="Find an asset"></div>
    <div class="picker">${pickHtml||'<div class="empty">No assets match.</div>'}</div>
    ${t.lines.length?`<div class="lines">${t.lines.map(l=>{const it=itemById(l.itemId); return `<div class="line"><div><b>#${esc(l.itemId)}</b> ${esc(it?.name||"")}<div class="calc">${esc(l.label)} · rates ${money(l.day)} / ${money(l.week)} / ${money(l.month)}</div></div><div class="num"><b>${money(l.total)}</b></div><button class="del" data-act="unpick" data-id="${esc(l.itemId)}" aria-label="Remove">×</button></div>`;}).join("")}</div>`:`<div class="muted small">No equipment added yet.</div>`}
    <div class="subhead"><span>Delivery</span></div>
    <div class="form">
      <div class="field"><label for="f-dzone">Zone</label><select id="f-dzone" data-act-change="dzone"><option value="">No delivery</option>${S.settings.delivery.map(d=>`<option value="${esc(d.name)}" ${o.delivery?.name===d.name?"selected":""}>${esc(d.name)} — ${money(d.rate)}/trip</option>`).join("")}</select></div>
      <div class="field"><label for="f-trips">Trips</label><input id="f-trips" type="number" min="0" data-of="delivery.trips" value="${o.delivery?.trips??0}"><span class="hint">2 = drop off and pick up</span></div>
      <div class="field"><label for="f-dtotal">Delivery total</label><input id="f-dtotal" disabled value="${money(t.delivery)}"></div>
    </div>
    <div class="subhead"><span>Inventory sold</span><button class="btn sm" data-act="add-sale" ${S.inventory.some(i=>!i.retired&&canSell(i))?"":"disabled"}>+ Add item</button></div>
    ${salesEditor(o.sales||[])}
    <div class="subhead"><span>Other charges</span><button class="btn sm" data-act="add-charge">+ Add charge</button></div>
    ${chargesEditor(o.charges)}
    <div class="form"><div class="field"><label for="f-adj">Adjustment ($)</label><input id="f-adj" type="number" step="0.01" data-of="adjustment" value="${o.adjustment||0}"><span class="hint">Use a negative number for a credit</span></div>
    <div class="field" style="grid-column:span 2"><label for="f-notes">Internal notes</label><input id="f-notes" data-of="notes" value="${esc(o.notes)}"></div></div>
  </div>
  <footer><span class="total num">Total ${money(t.total)}</span><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="save-order" ${valid&&(o.lines.length||(o.sales||[]).length)?"":"disabled"}>${o.no?(o.status==="Booked"&&S.orders.find(x=>x.no===o.no)?.status!=="Booked"?"Save as booked":o.status==="Rented Out"&&S.orders.find(x=>x.no===o.no)?.status==="Booked"?"Rent out now":"Save changes"):o.status==="Booked"?"Book order":"Rent out now"}</button></footer>`);
}
function salesEditor(sl){
  const inv = S.inventory.filter(i=>!i.retired && canSell(i));
  if (!sl.length) return `<div class="muted small">${inv.length?"Nothing sold on this order. Add fuel, consumables or other inventory to bill it on this rental. Stock comes off when the order goes out.":"No inventory items are set to <b>Can be sold</b>. Turn it on when editing an item on the Inventory page."}</div>`;
  return `<div style="display:flex;flex-direction:column;gap:8px">${sl.map((x,i)=>{ const it = invById(x.invId), opts = inv.concat(it&&!inv.includes(it)?[it]:[]);
    const avail = it ? +(it.qty + ((draft?.salesTaken||{})[x.invId]||0)).toFixed(2) : 0, short = it && +x.qty > avail;
    return `<div class="charge-row">
    <select id="sl-n-${i}" data-sl="${i}" data-f="invId" aria-label="Inventory item">${opts.map(v=>`<option value="${v.id}" ${v.id===x.invId?"selected":""}>${esc(v.name)} · ${v.id===x.invId?avail:v.qty} ${esc(v.unit||"")} on hand</option>`).join("")}</select>
    <input id="sl-q-${i}" type="number" step="0.1" min="0" data-sl="${i}" data-f="qty" value="${x.qty}" aria-label="Quantity${it?` (${esc(it.unit||"")})`:""}">
    <input id="sl-r-${i}" type="number" step="0.01" min="0" data-sl="${i}" data-f="price" value="${x.price}" aria-label="Price each">
    <span class="num small amt" style="text-align:right">${money((+x.qty||0)*(+x.price||0))}</span>
    <button class="del" data-act="del-sale" data-i="${i}" aria-label="Remove">×</button></div>${short?`<div class="small" style="color:var(--warn)">Only ${avail} ${esc(it.unit||"")} of ${esc(it.name)} on hand, so stock will show below zero until you receive more.</div>`:""}`; }).join("")}</div>`;
}
function chargesEditor(ch){
  if (!ch.length) return `<div class="muted small">No extra charges.</div>`;
  return `<div style="display:flex;flex-direction:column;gap:8px">${ch.map((c,i)=>`<div class="charge-row">
    <select id="ch-n-${i}" data-ch="${i}" data-f="name" aria-label="Charge">${S.settings.charges.map(s=>`<option ${s.name===c.name?"selected":""}>${esc(s.name)}</option>`).join("")}${S.settings.charges.some(s=>s.name===c.name)?"":`<option selected>${esc(c.name)}</option>`}</select>
    <input id="ch-q-${i}" type="number" step="0.1" data-ch="${i}" data-f="qty" value="${c.qty}" aria-label="Quantity">
    <input id="ch-r-${i}" type="number" step="0.01" data-ch="${i}" data-f="rate" value="${c.rate}" aria-label="Rate">
    <span class="num small amt" style="text-align:right">${money((+c.qty||0)*(+c.rate||0))}</span>
    <button class="del" data-act="del-charge" data-i="${i}" aria-label="Remove">×</button></div>`).join("")}</div>`;
}
function nextFreeOrderNo(){ while (S.orders.some(x=>x.no===S.nextOrder)) S.nextOrder++; return S.nextOrder; }
function saveOrder(){
  const o = draft, at = toLocalInput(now()), origNo = o.no;
  let want = null;
  if (me()?.role==="Admin" && o.newNo!=null && String(o.newNo).trim()!==""){
    want = +o.newNo;
    if (!Number.isInteger(want) || want < 1 || want > 99999999){ toast("Order # has to be a whole number, 1 or higher"); return; }
    if (S.orders.some(x=>x.no===want && x.no!==origNo)){ toast(`Order #${want} already exists. Pick another number.`); return; }
  }
  delete o.pq; delete o.newNo; if (!o.cycle) delete o.cycle;
  if (!origNo){
    o.no = want ?? nextFreeOrderNo(); o.created = at; o.returnedAt = null; o.log=[{by:myName(),at,text:`Order created as ${o.status}`}];
    S.orders.push(o); if (o.no >= S.nextOrder) S.nextOrder = o.no + 1; toast(`Order #${o.no} created`);
  } else {
    const i = S.orders.findIndex(x=>x.no===origNo); (o.log ||= []).push({by:myName(),at,text:"Order edited"});
    if (want!=null && want!==origNo){ o.no = want; o.log.push({by:myName(),at,text:`Order number changed from #${origNo} to #${want}`}); ui.orders.sel.delete(origNo); }
    if (o.status==="Returned" && !o.returnedAt) o.returnedAt = o.ret; S.orders[i] = o;
    toast(o.no!==origNo ? `Order #${origNo} is now #${o.no}` : `Order #${o.no} saved`);
  }
  o.sales = (o.sales||[]).filter(x=>x.invId && +x.qty>0); if (!o.sales.length) delete o.sales;
  syncSales(o, "order edited");
  if (o.fromRequest){ const r = (S.requests||[]).find(x=>x.id===o.fromRequest); if (r && r.status==="Pending"){ reqDecide(r, "Approved", "", o.no); o.requestNo = r.no; toast(`${rqNo(r)} approved as order #${o.no}`); } delete o.fromRequest; }
  draft = null; save(); closeModal(); rerender();
}

/* ---- Order detail ---- */
function viewOrder(no){
  const o = S.orders.find(x=>x.no===no); if (!o) return;
  const c = custById(o.customerId), l = locById(o.locationId), t = orderTotals(o), st = orderStatus(o);
  const late = isOverdue(o) ? Math.floor((now()-parseD(o.ret))/DAY) : 0;
  const accrued = isOverdue(o) ? o.lines.reduce((s,x)=>s+lineRental(o,x,toLocalInput(now())).total,0) : 0;
  openModal(`<header><h2>Order #${o.no}</h2>${chipFor(st)}<button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">
    ${late?`<div class="warnbox"><b>${late} day${late>1?"s":""} past the return date.</b> If checked in today, rental comes to about <b class="num">${money(accrued)}</b> instead of ${money(t.rental)}.</div>`:""}
    <div class="kv">
      <div><span>Customer</span>${esc(c?.name||"—")}</div><div><span>Location</span>${esc(locLabel(l))}</div><div><span>Pickup</span>${esc(o.pickup||"--")}</div>
      <div><span>Rent out</span><span class="num" style="text-transform:none;letter-spacing:0;font-size:inherit;color:inherit;font-weight:400">${fmtDT(o.out).replace("<br>"," ")}</span></div>
      <div><span>${o.status==="Returned"?"Returned":"Return date"}</span><span class="num" style="text-transform:none;letter-spacing:0;font-size:inherit;color:inherit;font-weight:400">${fmtDT(orderRentalEnd(o)).replace("<br>"," ")}</span></div>
      <div><span>Description</span>${esc(o.description)||"—"}</div>
    </div>
    <div class="lines">${t.lines.map(x=>{const it=itemById(x.itemId); return `<div class="line"><div><button class="linkbtn" data-act="view-item" data-id="${esc(x.itemId)}"><b>#${esc(x.itemId)}</b></button> ${esc(it?.name||"(deleted asset)")}<div class="calc">${esc(x.label)} (${x.days} days)</div></div><div class="num">${money(x.total)}</div><span></span></div>`;}).join("")}
      ${t.delivery?`<div class="line"><div>Delivery · ${esc(o.delivery.name)} × ${o.delivery.trips}</div><div class="num">${money(t.delivery)}</div><span></span></div>`:""}
      ${(o.sales||[]).map(x=>{ const it = invById(x.invId); return `<div class="line"><div>${esc(it?.name||"(inventory item removed)")} <span class="chip neutral">Sold</span><div class="calc">${x.qty} ${esc(it?.unit||"")} × ${money(x.price)}</div></div><div class="num">${money(x.qty*x.price)}</div><span></span></div>`; }).join("")}
      ${(o.charges||[]).map(ch=>`<div class="line"><div>${esc(ch.name)} · ${ch.qty} × ${money(ch.rate)}</div><div class="num">${money(ch.qty*ch.rate)}</div><span></span></div>`).join("")}
      ${t.bundleDisc?`<div class="line"><div>Bundle discount · ${esc(o.bundle.name)} (${o.bundle.pct}%)</div><div class="num">-${money(t.bundleDisc)}</div><span></span></div>`:""}
      ${+o.adjustment?`<div class="line"><div>Adjustment</div><div class="num">${money(+o.adjustment)}</div><span></span></div>`:""}
    </div>
    ${cyclePanel(o)}
    ${o.notes?`<div class="small"><b>Notes:</b> ${esc(o.notes)}</div>`:""}
    ${o.returnNotes?`<div class="small"><b>Return condition:</b> ${esc(o.returnNotes)}</div>`:""}
    <div class="small muted">${(o.log||[]).map(e=>`${fmtDate(e.at)} · ${esc(e.text)}${e.by?" · "+esc(e.by):""}`).join("<br>")}</div>
  </div>
  <footer><span class="total num">Total ${money(t.total)}</span>
    ${ACTIVE.includes(o.status)?`<button class="btn danger" data-act="cancel-order" data-no="${o.no}">Cancel order</button>`:""}
    <button class="btn" data-act="dup-order" data-no="${o.no}">Duplicate</button>
    ${o.status==="Returned"||(o.status==="Rented Out"&&!o.cycle)?`<button class="btn" data-act="inv-open" data-no="${o.no}">${ico("printer")} ${o.status==="Returned"?(o.cycle?"Final invoice":"Invoice"):"Print invoice"}</button>`:""}
    <button class="btn" data-act="edit-order" data-no="${o.no}">Edit</button>
    ${o.status==="Rented Out"&&o.pickup!=="BHESC Pickup"?`<button class="btn" data-act="req-pickup" data-no="${o.no}">Set BHESC Pickup</button>`:""}
    ${o.status==="Booked"?`<button class="btn primary" data-act="mark-out" data-no="${o.no}">Rent out now</button>`:""}
    ${o.status==="Rented Out"?`<button class="btn primary" data-act="checkin" data-no="${o.no}">Check in</button>`:""}
  </footer>`);
}

/* ---- Check-in ---- */
let ci = null;
function openCheckin(no){
  const o = S.orders.find(x=>x.no===no);
  const loc = {}; o.lines.forEach(l=>{ const it = itemById(l.itemId); if (it) loc[l.itemId] = it.locationId || yardId(); });
  ci = {no, at:toLocalInput(now()), billActual:true, charges:[], notes:"", maint:{}, loc, pickup:(o.pickup==="BHESC Pickup"||o.pickup==="BHESC Delivered")?"BHESC Pickup":"Customer Returned"};
  drawCheckin(o);
}
function drawCheckin(o){
  o = o || S.orders.find(x=>x.no===ci.no);
  const end = ci.billActual ? ci.at : o.ret;
  const rental = o.lines.reduce((s,l)=>s+lineRental(o,l,end).total,0);
  const extra = ci.charges.reduce((s,c)=>s+(+c.qty||0)*(+c.rate||0),0);
  const base = orderTotals(o);
  const total = rental + base.delivery + base.charges + (base.sales||0) + extra + (+o.adjustment||0);
  openModal(`<header><h2>Check in order #${o.no}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">
    <div class="form">
      <div class="field"><label for="ci-pickup">How it came back</label><select id="ci-pickup" data-ci="pickup"><option ${ci.pickup==="Customer Returned"?"selected":""}>Customer Returned</option><option ${ci.pickup==="BHESC Pickup"?"selected":""}>BHESC Pickup</option></select></div>
      <div class="field"><label for="ci-at">Returned at</label><input id="ci-at" type="datetime-local" data-ci="at" value="${esc(ci.at)}"></div>
      <div class="field"><label for="ci-bill">Bill rental through</label><select id="ci-bill" data-ci="billActual"><option value="1" ${ci.billActual?"selected":""}>Actual return date</option><option value="0" ${ci.billActual?"":"selected"}>Original return date (${fmtDate(o.ret)})</option></select></div>
      <div class="field"><label>Rental</label><input disabled value="${money(rental)}" aria-label="Rental"></div>
    </div>
    <div class="subhead"><span>Return charges</span><span style="display:flex;gap:6px;flex-wrap:wrap">${["Diesel on return","Gas on return","Excess dirt","Concrete splatter","Mechanic service"].filter(n=>S.settings.charges.some(c=>c.name===n)).map(n=>`<button class="btn sm" data-act="ci-quick" data-v="${esc(n)}">+ ${esc(n.replace(" on return",""))}</button>`).join("")}</span></div>
    ${chargesEditor(ci.charges)}
    <div class="field"><label for="ci-notes">Condition notes</label><textarea id="ci-notes" data-ci="notes" placeholder="Damage, cleanliness, hours, missing parts">${esc(ci.notes)}</textarea></div>
    <div class="subhead"><span>Returning to</span><span class="small" style="text-transform:none;letter-spacing:0">Where each asset is kept from now on</span></div>
    <div class="lines">${o.lines.filter(l=>itemById(l.itemId)).map(l=>`<div class="line ret-line"><div><b>#${esc(l.itemId)}</b> ${esc(itemById(l.itemId)?.name||"")}<label class="small" style="display:flex;gap:6px;align-items:center;margin-top:4px"><input type="checkbox" data-act="ci-maint" data-id="${esc(l.itemId)}" ${ci.maint[l.itemId]?"checked":""} style="accent-color:var(--accent)"> Needs service before the next rental</label></div><span></span>${returnLocSelect("data-ci-loc", l.itemId, ci.loc[l.itemId])}</div>`).join("")}</div>
  </div>
  ${o.bills?.length?`<div class="infobox">${o.bills.length} cycle invoice${o.bills.length>1?"s":""} already sent for ${money(o.bills.reduce((t,b)=>t+b.total,0))}. The final invoice bills only what's left.</div>`:""}
  <footer><span class="total num">Final ${money(total)}</span><button class="btn" data-act="close">Cancel</button><button class="btn" data-act="ci-save">Complete check-in</button><button class="btn primary" data-act="ci-save" data-print="1">${ico("printer")} Check in &amp; ${o.cycle?"final invoice":"print invoice"}</button></footer>`);
}
function saveCheckin(andPrint){
  const o = S.orders.find(x=>x.no===ci.no);
  o.status = "Returned"; o.returnedAt = ci.billActual ? ci.at : o.ret; o.actualReturn = ci.at;
  o.charges = (o.charges||[]).concat(ci.charges.filter(c=>+c.qty)); o.returnNotes = ci.notes; o.pickup = ci.pickup || "Customer Returned";
  (o.log ||= []).push({by:myName(),at:ci.at,text:"Checked in"});
  Object.entries(ci.loc||{}).forEach(([id,loc])=>applyReturnLocation(id, loc, `order #${o.no}`)); Object.keys(ci.maint).filter(k=>ci.maint[k]).forEach(id=>newSR({itemId:id,date:ci.at.slice(0,10),type:"Corrective maintenance",severity:"Medium",opStatus:"Degraded",stage:"Accepted",desc:`From return of order #${o.no}${ci.notes?": "+ci.notes:""}`,why:`returned on order #${o.no}`}));
  ci = null; save(); closeModal(); toast(`Order #${o.no} checked in`); rerender(); if (andPrint) openInvoice(o.no);
}

/* ---- Item modals ---- */
/* ================= Asset page (photos, details, comments, files) ================= */
ui.asset = {tab:"details", bottom:"comments", closed:{}};
const SAFE_ID = /^[A-Za-z0-9-]+$/;
function openAsset(id){ closeModal(); $("#pmenu")?.remove(); ui.asset.tab = "details"; if (SAFE_ID.test(id) && !VIEWS["asset-"+id]) go("asset-"+id); else { ui.assetId = id; go("asset"); } }
function assetIdFromRoute(r){ return r==="asset" ? ui.assetId : r.slice(6); }
function ageText(fromStr){
  const a = parseD(fromStr+"T00:00"), b = now(); if (!a || a > b) return "—";
  let y = b.getFullYear()-a.getFullYear(), m = b.getMonth()-a.getMonth(), d = b.getDate()-a.getDate();
  if (d < 0){ m--; d += new Date(b.getFullYear(), b.getMonth(), 0).getDate(); } if (m < 0){ y--; m += 12; }
  const p = (n,w) => n ? `${n} ${w}${n>1?"s":""}` : "";
  return [p(y,"Year"),p(m,"Month"),p(d,"Day")].filter(Boolean).join(" ") || "0 Days";
}
function bookValue(it){
  const cost = +it.cost; if (!it.cost || !it.purchased) return null;
  const life = +it.lifeYears || 5, salv = +it.salvage || 0;
  const yrs = Math.max(0, (now() - parseD(it.purchased+"T00:00")) / (365.25*DAY));
  return Math.max(salv, cost - (cost - salv) * Math.min(1, yrs/life));
}
function itemBilled(id){ return S.orders.filter(o=>o.status!=="Cancelled").reduce((s,o)=>s+orderTotals(o).lines.filter(l=>l.itemId===id).reduce((a,l)=>a+l.total,0),0); }
function itemLog(it, text){ (it.log ||= []).push({at:toLocalInput(now()), by:myName(), text}); }
function primaryPhoto(it){ return (it.photos||[]).find(p=>p.primary) || (it.photos||[])[0]; }
function thumb(it, size=36){ const p = primaryPhoto(it); return p ? `<img class="thumb" src="${p.data}" alt="" width="${size}" height="${size}" style="width:${size}px;height:${size}px">` : `<span class="thumb ph" style="width:${size}px;height:${size}px">${ico(ICON_FOR[it.category]||"items",1.6)}</span>`; }
const ICON_FOR = {};
function avatarHtml(u, cls=""){ if (!u) return `<div class="avatar ${cls}"></div>`; return `<div class="avatar ${cls}">${u.photo?`<img src="${u.photo}" alt="">`:esc(initials(u.name))}</div>`; }
function userByName(n){ return S.users.find(u=>u.name===n); }
function fmtBytes(n){ return n>1048576 ? (n/1048576).toFixed(1)+" MB" : Math.max(1,Math.round(n/1024))+" KB"; }

function assetEvents(it){
  const ev = [], id = it.id;
  (it.log||[]).forEach(l=>ev.push({at:l.at,by:l.by,text:l.text,k:"Asset"}));
  S.orders.filter(o=>o.lines.some(l=>l.itemId===id)).forEach(o=>{
    ev.push({at:o.out,text:`${o.status==="Booked"?"Booked":"Rented out"} on order #${o.no} to ${locLabel(locById(o.locationId))}`,k:"Rental",no:o.no});
    if (o.status==="Returned") ev.push({at:o.actualReturn||o.returnedAt||o.ret,text:`Returned from order #${o.no}`,k:"Rental",no:o.no});
    if (o.status==="Cancelled") ev.push({at:(o.log||[]).slice(-1)[0]?.at||o.out,text:`Order #${o.no} cancelled`,k:"Rental",no:o.no}); });
  S.custody.filter(c=>c.kind==="asset"&&c.itemId===id).forEach(c=>{ ev.push({at:c.outAt,by:c.by,text:`Checked out to ${c.holder} (${c.ref})`,k:"Check-out"}); if (c.inAt) ev.push({at:c.inAt,by:c.inBy,text:`Checked in from ${c.holder}${c.inNotes?": "+c.inNotes:""}`,k:"Check-out"}); });
  S.maintenance.filter(m=>m.itemId===id).forEach(m=>{ (m.files||[]).forEach(f=>ev.push({at:f.at,by:f.by,text:`Service document added: ${f.name}`,k:"Service"})); ev.push({at:m.date+"T08:00",text:`${m.type} work order opened${m.desc?": "+m.desc:""}`,k:"Service"}); if (m.status==="Completed") ev.push({at:(m.doneAt||m.date)+"T17:00",text:`${m.type} work order completed`,k:"Service"}); });
  (it.comments||[]).forEach(c=>ev.push({at:c.at,by:c.by,text:"Comment: "+c.text,k:"Comment"}));
  (it.files||[]).forEach(f=>ev.push({at:f.at,by:f.by,text:"File added: "+f.name,k:"File"}));
  return ev.sort((a,b)=>b.at.localeCompare(a.at));
}

function vAsset(id){
  const it = itemById(id);
  if (!it) return band("Asset not found","") + `<div class="content"><div class="empty"><strong>No asset #${esc(id)}</strong>It may have been deleted or renumbered. <button class="linkbtn" data-act="goto" data-to="items">Back to assets</button></div></div>`;
  const st = itemStatus(it), o = activeOrderFor(id), cu = openCustody(id), res = nextReservation(id), mo = openMaint(id);
  const idx = S.items.indexOf(it), prev = S.items[idx-1], next = S.items[idx+1];
  const orders = S.orders.filter(x=>x.lines.some(l=>l.itemId===id)).sort((a,b)=>b.out.localeCompare(a.out));
  const svc = S.maintenance.filter(m=>m.itemId===id).sort((a,b)=>b.date.localeCompare(a.date));
  const cos = S.custody.filter(c=>c.kind==="asset"&&c.itemId===id).sort((a,b)=>b.outAt.localeCompare(a.outAt));
  const events = assetEvents(it);
  const lastSvc = svc.find(m=>m.status==="Completed");
  let line = "";
  if (o) { const c = custById(o.customerId); line = `Rented out at <button class="linkbtn" data-act="loc-orders" data-id="${o.locationId}">${esc(locLabel(locById(o.locationId)))}</button>${c?` (${esc(c.name)})`:""} through <button class="linkbtn" data-act="view-order" data-no="${o.no}">Order #${o.no}</button> · ${isOverdue(o)?`<b style="color:var(--bad)">due back ${fmtDate(o.ret)}</b>`:`Available on ${fmtDate(o.ret)} ${o.ret.slice(11)}`}`; }
  else if (cu) line = `Checked out to <b>${esc(cu.holder)}</b> at ${esc(locLabel(locById(cu.locationId)))} (${esc(cu.ref)})${cu.dueAt?` · due back ${fmtDate(cu.dueAt)}`:""}`;
  else if (mo) line = `In the shop on <button class="linkbtn" data-act="view-sr" data-id="${mo.id}">${srNo(mo)}</button>: ${esc(mo.type)} · ${esc(mo.stage)} since ${fmtDate((mo.acceptedAt||mo.date).slice(0,10)+"T00:00")}`;
  else if (openSR(id)) { const r = openSR(id); line = `Problem reported on <button class="linkbtn" data-act="view-sr" data-id="${r.id}">${srNo(r)}</button> (${esc(r.opStatus)}). Still in service until it's accepted into the shop`; }
  else if (st==="Retired") line = `Retired${it.retiredAt?` on ${fmtDate(it.retiredAt)}`:""}${it.retiredReason?` · ${esc(it.retiredReason)}`:""}${it.retiredBy?` · by ${esc(it.retiredBy)}`:""}${it.retiredNote?` · ${esc(it.retiredNote)}`:""}`;
  else line = `At ${esc(homeLabel(it))}${res?` · booked on <button class="linkbtn" data-act="view-order" data-no="${res.no}">#${res.no}</button> from ${fmtDate(res.out)}`:""}`;
  const kv = (label, val) => `<div class="kv-row"><dt>${label}</dt><dd>${val===""||val==null?'<span class="muted">—</span>':val}</dd></div>`;
  const sec = (key, title, body) => `<details class="sec" ${ui.asset.closed[key]?"":"open"} data-sec="${key}"><summary>${title}</summary><dl class="kvgrid">${body}</dl></details>`;
  const bv = bookValue(it);
  const u90 = utilization(id), daysOut = orders.filter(x=>["Rented Out","Returned"].includes(x.status)).reduce((s,x)=>s+rentalDays(x.out, x.status==="Returned"?(x.actualReturn||x.returnedAt||x.ret):toLocalInput(now())),0);
  const T = ui.asset.tab;
  let tabBody = "";
  if (T==="details") tabBody = `
    ${it.notes?`<div class="asset-notes">${esc(it.notes).replace(/\n/g,"<br>")}</div>`:""}
    ${sec("core","Core information",
      kv("Asset #", `<b>${esc(it.id)}</b>`) +
      kv("Product model number", esc(it.model)) + kv("Location", o?`<button class="linkbtn" data-act="loc-orders" data-id="${o.locationId}">${esc(locLabel(locById(o.locationId)))}</button>`:cu?esc(locLabel(locById(cu.locationId))):esc(homeLabel(it))) + (o||cu?kv("Home location", esc(homeLabel(it))):"") +
      kv("Group", (()=>{ const g = groupOf(it.category); return (g?.parent&&groupOf(g.parent)?`<button class="linkbtn" data-act="group-open" data-v="${esc(g.parent)}">${esc(g.parent)}</button> › `:"")+`<button class="linkbtn" data-act="group-assets" data-v="${esc(it.category)}">${esc(it.category)}</button>`; })()) + kv("Rented out at", o?`<button class="linkbtn" data-act="view-order" data-no="${o.no}">#${o.no} · ${esc(locLabel(locById(o.locationId)))}</button>`:"") +
      kv("Last service date", lastSvc?fmtDate((lastSvc.doneAt||lastSvc.date)+"T00:00"):"") + kv("Upcoming service date", it.nextService?`<span class="num" ${it.nextService<toLocalInput(now()).slice(0,10)?'style="color:var(--bad);font-weight:600"':""}>${fmtDate(it.nextService+"T00:00")}</span>`:"") +
      kv("Check-in due", o?`<span class="num">${fmtDT(o.ret).replace("<br>"," ")}</span>`:cu&&cu.dueAt?`<span class="num">${fmtDT(cu.dueAt).replace("<br>"," ")}</span>`:"") +
      kv("Hours", it.hours!=null&&it.hours!==""?`<span class="num">${Number(it.hours).toLocaleString()}</span> <button class="linkbtn small" data-act="a-hours" data-id="${esc(id)}">Update</button>`:`<button class="linkbtn small" data-act="a-hours" data-id="${esc(id)}">Add reading</button>`) +
      kv("Serial number", esc(it.serial)) + kv("Status", chipFor(st)) + kv("Service interval", it.svcHours?`Every ${it.svcHours} hours`:""))}
    ${sec("purchase","Purchase information",
      kv("Purchased on", it.purchased?fmtDate(it.purchased+"T00:00"):"") + kv("Cost price", it.cost?`<span class="num">${money(it.cost)}</span>`:"") +
      kv("Vendor", esc(it.vendor)) + kv("Age", it.purchased?ageText(it.purchased):"") +
      kv("Added to fleet", it.created?fmtDate(it.created):"") + kv("Depreciated value", bv!=null?`<span class="num">${money(bv)}</span> <span class="muted small">(straight line, ${+it.lifeYears||5} yr life${+it.salvage?`, ${money(it.salvage)} salvage`:""})</span>`:"") +
      kv("Item-level tax rate", it.taxRate?esc(it.taxRate)+"%":"") + kv("Warranty until", it.warranty?fmtDate(it.warranty+"T00:00"):""))}
    ${sec("additional","Additional information",
      kv("Manufacturer", esc(it.make)) + kv("Model year", esc(it.year)) + kv("Engine / fuel", esc(it.fuel)) + kv("Weight", esc(it.weight)))}
    ${sec("pricing","Pricing",
      kv("Rental pricing", `<span class="num">${it.day!=null?money(it.day)+" per day<br>":""}${it.week!=null?money(it.week)+" per week<br>":""}${it.month!=null?money(it.month)+" per month":""}</span>`) +
      kv("Rental billed to date", `<span class="num">${money(itemBilled(id))}</span>`) + kv("Damage fee", it.damageFee!=null&&it.damageFee!==""?`<span class="num">${money(it.damageFee)}</span>`:"") +
      kv("Sale price", it.salePrice!=null&&it.salePrice!==""?`<span class="num">${money(it.salePrice)}</span>`:"") +
      kv("Return on cost", it.cost?`<span class="num">${Math.round(itemBilled(id)/it.cost*100)}%</span> <span class="muted small">of purchase price billed</span>`:""))}
    ${sec("usage","Usage information",
      kv("Time on rent, 90 days", `<span class="num">${Math.round(u90*100)}%</span>`) + kv("Days rented, all time", `<span class="num">${daysOut}</span>`) +
      kv("Rental orders", `<span class="num">${orders.filter(x=>x.status!=="Cancelled").length}</span>`) + kv("Crew check-outs", `<span class="num">${cos.length}</span>`) +
      kv("Work orders", `<span class="num">${svc.length}</span>`) + kv("Last activity", lastUsed(id)?fmtDate(lastUsed(id)):""))}`;
  else if (T==="orders") tabBody = `<div class="tablewrap"><table><thead><tr><th>Order #</th><th>Job</th><th>Customer</th><th>Out</th><th>Back</th><th>Billing</th><th style="text-align:right">This asset</th><th>Status</th></tr></thead><tbody>${orders.length?orders.map(x=>{const ln=orderTotals(x).lines.find(l=>l.itemId===id); return `<tr><td><button class="linkbtn num" data-act="view-order" data-no="${x.no}">${x.no}</button></td><td>${esc(locLabel(locById(x.locationId)))}</td><td>${esc(custById(x.customerId)?.name||"")}</td><td class="num">${fmtDate(x.out)}</td><td class="num">${fmtDate(orderRentalEnd(x))}</td><td class="small muted">${esc(ln?.label||"")}</td><td class="num" style="text-align:right">${money(ln?.total)}</td><td>${chipFor(orderStatus(x))}</td></tr>`;}).join(""):`<tr><td colspan="8"><div class="empty">Never rented.</div></td></tr>`}</tbody></table></div>`;
  else if (T==="services") tabBody = `<div class="tablewrap"><table><thead><tr><th>SR #</th><th>Reported</th><th>Severity</th><th>Type</th><th>Problem</th><th>Mechanic</th><th style="text-align:right">Hours</th><th style="text-align:right">Cost</th><th>Docs</th><th>Status</th><th></th></tr></thead><tbody>${svc.length?svc.map(m=>`<tr><td class="num"><b>${srNo(m)}</b></td><td class="num">${fmtDate(m.date+"T00:00")}</td><td>${sevChip(m.severity)}</td><td>${esc(m.type)}</td><td>${esc(m.desc)}</td><td>${esc(m.tech)}</td><td class="num" style="text-align:right">${m.hours||"—"}</td><td class="num" style="text-align:right">${money(srCost(m))}</td><td>${docsBadge(m)||'<span class="muted">—</span>'}</td><td>${chipFor(m.stage||m.status)}</td><td><button class="btn sm" data-act="view-sr" data-id="${m.id}">Open</button></td></tr>`).join(""):`<tr><td colspan="11"><div class="empty">No service logged yet.</div></td></tr>`}</tbody></table></div>`;
  else if (T==="checkouts") tabBody = `<div class="tablewrap"><table><thead><tr><th>Ref</th><th>Assigned to</th><th>Job / location</th><th>Out</th><th>Back</th><th>Notes</th><th>Status</th></tr></thead><tbody>${cos.length?cos.map(c=>`<tr><td class="num"><b>${esc(c.ref)}</b></td><td>${esc(c.holder)}</td><td>${esc(locLabel(locById(c.locationId)))}</td><td class="num">${fmtDate(c.outAt)}</td><td class="num">${c.inAt?fmtDate(c.inAt):"—"}</td><td class="small">${esc(c.inNotes||c.outNotes||"")}</td><td>${chipFor(c.inAt?"Returned":custodyOverdue(c)?"Overdue":"Checked Out")}</td></tr>`).join(""):`<tr><td colspan="7"><div class="empty">Never checked out to a crew.</div></td></tr>`}</tbody></table></div>`;
  else tabBody = `<div class="timeline">${events.length?events.map(e=>`<div class="ev"><span class="ev-k">${esc(e.k)}</span><div><div>${e.no?`<button class="linkbtn" data-act="view-order" data-no="${e.no}">${esc(e.text)}</button>`:esc(e.text)}</div><div class="small muted num">${fmtDate(e.at)} ${e.at.length>10?e.at.slice(11,16):""}${e.by?" · "+esc(e.by):""}</div></div></div>`).join(""):'<div class="empty">No activity yet.</div>'}</div>`;

  const photosHtml = `<div class="dropzone" data-drop="photos" data-id="${esc(id)}">
      <div class="gallery">${(it.photos||[]).map((p,i)=>`<figure class="ph-card ${p.primary?"primary":""}"><button class="ph-img" data-act="a-photo-view" data-id="${esc(id)}" data-i="${i}" aria-label="View photo"><img src="${p.data}" alt="${esc(p.name||"")}"></button>
        <figcaption><span class="small muted">${fmtDate(p.at)}${p.by?" · "+esc(p.by):""}</span><span style="display:flex;gap:4px">${p.primary?'<span class="chip available">Main</span>':`<button class="btn sm" data-act="a-photo-primary" data-id="${esc(id)}" data-i="${i}">Make main</button>`}<button class="del" data-act="a-photo-del" data-id="${esc(id)}" data-i="${i}" aria-label="Delete photo">×</button></span></figcaption></figure>`).join("")}
        <button class="ph-add" data-act="a-photo-add" data-id="${esc(id)}">${ico("plus")}<span>Add photos</span><span class="small muted">or drop images here</span></button></div></div>`;
  if (ui.asset.tab==="photos") ui.asset.tab = "details";
  const B = ui.asset.bottom, comments = (it.comments||[]).slice().sort((a,b)=>b.at.localeCompare(a.at)), files = it.files||[];
  const canNotes = can("notes"), meU = me();
  const bottom = B==="comments" ? `
    ${canNotes?`<div class="cmt-new">${avatarHtml(meU,"sm")}<textarea id="a-cmt" placeholder="Add a comment. Damage found, parts on order, customer requests…" aria-label="New comment"></textarea><button class="btn primary" data-act="a-comment" data-id="${esc(id)}">Post</button></div>`:""}
    <div class="cmts">${comments.length?comments.map(c=>`<div class="cmt">${avatarHtml(userByName(c.by)||{name:c.by},"sm")}<div style="min-width:0;flex:1"><div class="cmt-meta"><b>${esc(c.by)}</b><span class="muted small num">${fmtDate(c.at)} ${c.at.slice(11,16)}</span>${(meU&&(c.by===meU.name||meU.role==="Admin"))?`<button class="linkbtn small" data-act="a-comment-del" data-id="${esc(id)}" data-c="${c.id}" style="margin-left:auto;color:var(--muted)">Delete</button>`:""}</div><div class="cmt-text">${esc(c.text).replace(/\n/g,"<br>")}</div></div></div>`).join(""):'<div class="muted small" style="padding:6px 0">No comments yet.</div>'}</div>`
  : `<div class="dropzone" data-drop="files" data-id="${esc(id)}">
    <div class="files">${files.map((f,i)=>`<div class="file">${f.type.startsWith("image/")?`<img src="${f.data}" alt="" class="file-ic">`:`<span class="file-ic ext">${esc((f.name.split(".").pop()||"file").slice(0,4).toUpperCase())}</span>`}<div style="min-width:0;flex:1"><button class="linkbtn trunc" style="max-width:100%" data-act="a-file-open" data-id="${esc(id)}" data-i="${i}">${esc(f.name)}</button><div class="small muted">${fmtBytes(f.size)} · ${fmtDate(f.at)}${f.by?" · "+esc(f.by):""}</div></div>${canNotes?`<button class="del" data-act="a-file-del" data-id="${esc(id)}" data-i="${i}" aria-label="Delete file">×</button>`:""}</div>`).join("")}
    ${canNotes?`<button class="ph-add file-add" data-act="a-file-add" data-id="${esc(id)}">${ico("plus")}<span>Attach files</span><span class="small muted">Manuals, invoices, inspection sheets, titles · up to 3 MB each · or drop here</span></button>`:""}
    ${!files.length&&!canNotes?'<div class="muted small">No files attached.</div>':""}</div></div>`;

  const ph = primaryPhoto(it);
  return `<div class="asset-head">
    <button class="asset-photo" data-act="${ph?"a-photo-view":"a-photo-add"}" data-id="${esc(id)}" data-i="${ph?it.photos.indexOf(ph):0}" aria-label="${ph?"View photo":"Add a photo"}">${ph?`<img src="${ph.data}" alt="">`:`${ico("camera",1.6)}<span>Add photo</span>`}</button>
    <div style="min-width:0;flex:1"><h1>#${esc(it.id)} - ${esc(it.name)}</h1><div class="asset-line">${line}</div></div>
    <div class="asset-nav"><button class="icon-sq" data-act="a-go" data-id="${esc(prev?.id||"")}" ${prev?"":"disabled"} aria-label="Previous asset" title="Previous asset">‹</button><button class="icon-sq" data-act="a-go" data-id="${esc(next?.id||"")}" ${next?"":"disabled"} aria-label="Next asset" title="Next asset">›</button><button class="icon-sq" data-act="goto" data-to="items" aria-label="Close" title="Back to assets">×</button></div>
  </div>
  <div class="asset-bar">
    ${chipFor(st)}
    <span class="bar-sep"></span>
    <button class="icon-sq" data-act="edit-item" data-id="${esc(id)}" title="Edit asset" aria-label="Edit asset">${ico("edit")}</button>
    <button class="icon-sq" data-act="dup-item" data-id="${esc(id)}" title="Duplicate asset" aria-label="Duplicate asset">${ico("copy")}</button>
    <button class="icon-sq" data-act="del-item" data-id="${esc(id)}" title="Delete asset" aria-label="Delete asset">${ico("trash")}</button>
    ${st==="Retired"?`<button class="btn sm" data-act="a-restore" data-id="${esc(id)}">Restore to fleet</button>`:`<button class="btn sm" data-act="a-retire" data-id="${esc(id)}">Retire</button>`}
    <div class="bar-right">
      ${st==="Available"?`<button class="btn" data-act="co-item" data-id="${esc(id)}">Check out</button>`:""}${cu?`<button class="btn" data-act="co-in" data-id="${cu.id}">Check in</button>`:""}
      ${st==="Retired"?"":`<button class="btn" data-act="new-maint" data-id="${esc(id)}">${ico("maintenance")} New service request</button>`}
      ${st==="Available"&&it.shopOnly?`<span class="chip shop" style="align-self:center">Shop use only</span>`:st==="Available"?`<button class="btn primary" data-act="rent-item" data-id="${esc(id)}">${ico("orders")} Add to order</button>`:o?`<button class="btn primary" data-act="view-order" data-no="${o.no}">Open order #${o.no}</button>`:""}
    </div>
  </div>
  ${tabs([["details","Details"],["orders","Orders",orders.length],["services","Services",svc.length],["checkouts","Check-outs",cos.length],["events","Events",events.length]],T,"atab")}
  <div class="content asset-body">${tabBody}
    <section class="panel bottom-panel">${tabs([["comments","Comments",(it.comments||[]).length],["files","Files",files.length],["photos","Photos",(it.photos||[]).length]],B,"abtab")}<div class="body">${B==="photos"?photosHtml:bottom}</div></section>
  </div>`;
}

/* ---- Asset form (all fields) ---- */
function itemForm(id, copyOf){
  const src = id ? itemById(id) : copyOf ? itemById(copyOf) : null;
  const it = src ? {...src} : {id:"",name:"",category:S.groups[0]?.name||"Site Services",day:"",week:"",month:"",status:"Active",serial:"",notes:""};
  if (copyOf){ let n = 2; while (itemById(`${copyOf}-${n}`)) n++; it.id = `${copyOf}-${n}`; it.serial=""; it.tag=""; it.gpsId=""; it.hours=""; }
  const cats = S.groups.map(g=>g.name).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  const f = (k, label, type="text", extra="") => `<div class="field"><label for="it-${k}">${label}</label><input id="it-${k}" name="${k}" type="${type}" ${type==="number"?'step="0.01"':""} value="${esc(it[k]??"")}" ${extra}></div>`;
  openModal(`<header><h2>${id?`Edit #${esc(id)}`:copyOf?`Duplicate #${esc(copyOf)}`:"Add asset"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="itemform" data-orig="${esc(id||"")}" style="display:flex;flex-direction:column;gap:16px">
    <div class="subhead"><span>Core information</span></div>
    <div class="form">
      ${f("id","Asset #","text","required")}
      <div class="field" style="grid-column:span 2"><label for="it-name">Description</label><input id="it-name" name="name" required value="${esc(it.name)}"></div>
      <div class="field"><label for="it-cat">Group</label><select id="it-cat" name="category">${groupOptions(it.category)}</select></div>
      ${f("model","Product model number")}${f("serial","Serial number / VIN")}
      <div class="field"><label for="it-loc">Location</label><select id="it-loc" name="locationId">${sortedLocs().map(l=>`<option value="${l.id}" ${l.id===(it.locationId||yardId())?"selected":""}>${esc(locOpt(l))}</option>`).join("")}</select>${(()=>{ if (!id) return ""; const o = activeOrderFor(id), cu = openCustody(id); return o ? `<span class="hint">On rental at ${esc(locLabel(locById(o.locationId)))} (#${o.no}). This is where it goes when it comes back.</span>` : cu ? `<span class="hint">Checked out to ${esc(cu.holder)}. This is where it goes when it comes back.</span>` : ""; })()}</div>
      ${f("home","Spot / bay (optional)","text",'placeholder="Bay 3, north fence"')}
      ${f("hours","Hours (meter)","number")}${f("svcHours","Service every (hours)","number")}${f("nextService","Upcoming service date","date")}

      <div class="field full"><label class="tgl" for="it-rent"><input id="it-rent" type="checkbox" ${it.shopOnly?"":"checked"}><span class="tgl-ui" aria-hidden="true"></span><span><b>Available to rent</b><span class="hint" style="display:block">Off: shop use only. It won't show to superintendents, can't go on a rental order and isn't counted in the rental fleet. It can still be checked out to crews and serviced.</span></span></label></div>
      <div class="field full"><label for="it-notes">Notes shown at the top of the asset</label><textarea id="it-notes" name="notes" placeholder="Bucket PN: 279-5372&#10;Bucket SN: …">${esc(it.notes)}</textarea></div>
    </div>
    <div class="subhead"><span>Purchase information</span></div>
    <div class="form">${f("purchased","Purchased on","date")}${f("vendor","Vendor")}${f("cost","Cost price","number")}${f("lifeYears","Useful life (years)","number",'placeholder="5"')}${f("salvage","Salvage value","number")}${f("taxRate","Item-level tax rate %","number")}${f("warranty","Warranty until","date")}</div>
    <div class="subhead"><span>Additional information</span></div>
    <div class="form">${f("make","Manufacturer")}${f("year","Model year")}${f("fuel","Engine / fuel")}${f("weight","Operating weight")}</div>
    <div class="subhead"><span>Pricing</span></div>
    <div class="form">${f("day","Day rate","number")}${f("week","Week rate","number")}${f("month","Month rate","number")}${f("damageFee","Damage fee","number")}${f("salePrice","Sale price","number")}</div>
    <span class="hint">Leave a rate blank if it isn't offered. Rate changes apply to new orders only.</span>
  </form></div>
  <footer>${id?`<button class="btn danger" data-act="del-item" data-id="${esc(id)}" style="margin-right:auto">Delete</button>`:""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="save-item">Save asset</button></footer>`);
}
function saveItem(){
  const f = $("#itemform"); const fd = Object.fromEntries(new FormData(f)); const orig = f.dataset.orig;
  fd.id = fd.id.trim(); fd.name = fd.name.trim(); if (!("status" in fd)) fd.status = orig ? (itemById(orig).status||"Active") : "Active";
  if (!fd.id || !fd.name) { toast("Asset # and description are required"); return; }
  if (fd.id !== orig && itemById(fd.id)) { toast(`Asset #${fd.id} already exists`); return; }
  const n = v => v===""||v==null?null:+v;
  ["day","week","month","cost","salvage","lifeYears","taxRate","damageFee","salePrice","hours","svcHours"].forEach(k=>fd[k]=n(fd[k]));
  fd.shopOnly = !$("#it-rent")?.checked;
  if (fd.shopOnly && orig && S.orders.some(o=>ACTIVE.includes(o.status) && o.lines.some(l=>l.itemId===orig))){ toast(`#${orig} is on an open rental order. Return or cancel it first, then make it shop use only.`); return; }
  if (orig){
    const it = itemById(orig), changes = [];
    if (it.hours !== fd.hours && fd.hours!=null) changes.push(`Hours ${it.hours??"—"} → ${fd.hours}`);
    ["day","week","month"].forEach(k=>{ if (it[k] !== fd[k]) changes.push(`${k[0].toUpperCase()+k.slice(1)} rate ${money(it[k])} → ${money(fd[k])}`); });
    if (it.category !== fd.category) changes.push(`Group ${it.category} → ${fd.category}`);
    if ((it.locationId||yardId()) !== fd.locationId) changes.push(`Location ${locLabel(locById(it.locationId||yardId()))} → ${locLabel(locById(fd.locationId))}`);
    if (it.status !== fd.status) changes.push(`Fleet status ${it.status} → ${fd.status}`);
    if (!!it.shopOnly !== fd.shopOnly) changes.push(fd.shopOnly ? "Set to shop use only" : "Made available to rent");
    Object.assign(it, fd); if (!it.shopOnly) delete it.shopOnly; itemLog(it, changes.length ? "Edited: "+changes.join(", ") : "Details edited");
    if (orig !== fd.id){ S.orders.forEach(o=>o.lines.forEach(l=>{ if(l.itemId===orig) l.itemId=fd.id; })); S.maintenance.forEach(m=>{ if(m.itemId===orig) m.itemId=fd.id; }); S.custody.forEach(c=>{ if(c.itemId===orig) c.itemId=fd.id; }); S.bundles.forEach(b=>b.members=b.members.map(x=>x===orig?fd.id:x)); }
    save(); closeModal(); toast("Asset saved"); if (!VIEWS[route()] && orig!==fd.id) openAsset(fd.id); else rerender();
  } else {
    const rec = {...fd, created:toLocalInput(now()), photos:[], comments:[], files:[], log:[]}; if (!rec.shopOnly) delete rec.shopOnly; itemLog(rec, "Asset added to the fleet");
    S.items.push(rec); S.items.sort((a,b)=>a.id.localeCompare(b.id,undefined,{numeric:true}));
    save(); toast("Asset saved"); openAsset(rec.id);
  }
}

const RETIRE_REASONS = ["Sold","Scrapped","Lost or stolen","Transferred to another company","Worn out / end of life","Other"];
function retireForm(kind, id){
  const isA = kind==="asset", rec = isA ? itemById(id) : invById(id);
  if (isA){
    const o = activeOrderFor(id), cu = openCustody(id), res = nextReservation(id);
    if (o){ toast(`#${id} is on rental #${o.no}. Check it in before retiring it.`); return; }
    if (cu){ toast(`#${id} is checked out to ${cu.holder}. Check it in before retiring it.`); return; }
    if (res){ toast(`#${id} is booked on order #${res.no}. Remove it from that order first.`); return; }
  } else if (invOut(id) > 0){ toast(`${invOut(id)} ${rec.unit} of ${rec.name} is checked out. Check it in before retiring it.`); return; }
  const mo = isA ? S.maintenance.filter(m=>m.itemId===id && m.status==="Open") : [];
  openModal(`<header><h2>Retire ${isA?`#${esc(id)} ${esc(rec.name)}`:esc(rec.name)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><p style="margin:0">${isA?"Retired assets leave the rental pool and can't be rented, checked out or put on a work order. Their history, photos and files stay, and you can restore them later.":"Retired inventory is hidden from check-outs and work orders. Its history stays, and you can restore it later."}</p>
  ${mo.length?`<div class="warnbox">${mo.length} open work order${mo.length>1?"s":""} will be marked done.</div>`:""}
  <form id="retireform" class="form" data-kind="${kind}" data-id="${esc(id)}">
    <div class="field"><label for="rt-reason">Reason</label><select id="rt-reason" name="reason">${RETIRE_REASONS.map(x=>`<option>${x}</option>`).join("")}</select></div>
    <div class="field"><label for="rt-date">Retired on</label><input id="rt-date" type="date" name="date" value="${toLocalInput(now()).slice(0,10)}"></div>
    <div class="field full"><label for="rt-note">Notes</label><input id="rt-note" name="note" placeholder="${isA?"Sold to, sale price, auction lot…":"Why it's no longer stocked"}"></div>
  </form></div>
  <footer><button class="btn" data-act="close">Cancel</button><button class="btn danger" data-act="retire-save">Retire</button></footer>`,"narrow");
}
function saveRetire(){
  const f = $("#retireform"), fd = Object.fromEntries(new FormData(f)), isA = f.dataset.kind==="asset", id = f.dataset.id;
  const rec = isA ? itemById(id) : invById(id), at = toLocalInput(now());
  Object.assign(rec, {retiredAt:fd.date||at.slice(0,10), retiredReason:fd.reason, retiredNote:fd.note, retiredBy:myName()});
  if (isA){
    rec.status = "Retired";
    S.maintenance.filter(m=>m.itemId===id && m.status==="Open").forEach(m=>{ m.status="Completed"; m.stage="Closed"; m.doneAt=at.slice(0,10); m.opStatus="Operational"; srLog(m, "Closed when the asset was retired"); });
    itemLog(rec, `Retired: ${fd.reason}${fd.note?` (${fd.note})`:""}`);
  } else { rec.retired = true; (rec.log ||= []).push({by:myName(),at,delta:0,reason:`Retired: ${fd.reason}${fd.note?` (${fd.note})`:""}`}); }
  save(); closeModal(); rerender(); toast(`${isA?"#"+id:rec.name} retired`);
}
function restoreRetired(kind, id){
  const isA = kind==="asset", rec = isA ? itemById(id) : invById(id);
  ["retiredAt","retiredReason","retiredNote","retiredBy"].forEach(k=>delete rec[k]);
  if (isA){ rec.status = "Active"; itemLog(rec, "Restored to the fleet"); } else { delete rec.retired; (rec.log ||= []).push({by:myName(),at:toLocalInput(now()),delta:0,reason:"Restored"}); }
  save(); closeModal(); rerender(); toast(`${isA?"#"+id:rec.name} restored`);
}
function returnLocSelect(attr, key, current){
  return `<select class="ret-loc" ${attr}="${esc(key)}" aria-label="Return location">${sortedLocs().map(l=>`<option value="${l.id}" ${l.id===current?"selected":""}>${esc(locOpt(l))}${l.isYard?" (default)":""}</option>`).join("")}</select>`;
}
function applyReturnLocation(itemId, locId, why){
  const it = itemById(itemId); if (!it || !locId) return;
  const before = it.locationId || yardId();
  if (before !== locId){ it.locationId = locId; itemLog(it, `Returned to ${locLabel(locById(locId))} (was ${locLabel(locById(before))}) · ${why}`); }
}
function hoursForm(id){ const it = itemById(id);
  openModal(`<header><h2>Hours for #${esc(id)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><div class="form"><div class="field"><label for="hr-val">Meter reading</label><input id="hr-val" type="number" step="0.1" value="${it.hours??""}"></div><div class="field"><label for="hr-note">Note</label><input id="hr-note" placeholder="Read at return from 25AU013"></div></div>
  <div class="subhead"><span>Earlier readings</span></div><div class="lines">${(it.log||[]).filter(l=>/^Hours|Edited: .*Hours/.test(l.text)).slice(-6).reverse().map(l=>`<div class="line"><div>${esc(l.text)}<div class="calc">${fmtDate(l.at)}${l.by?" · "+esc(l.by):""}</div></div><span></span><span></span></div>`).join("")||'<div class="line muted">No earlier readings.</div>'}</div></div>
  <footer><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="a-hours-save" data-id="${esc(id)}">Save reading</button></footer>`,"narrow"); }
document.addEventListener("toggle", e => { const d = e.target; if (d.matches?.("details.sec")) ui.asset.closed[d.dataset.sec] = !d.open; }, true);
/* ---- Photos & files ---- */
function pickFiles(accept, multiple, cb){ const inp = document.createElement("input"); inp.type = "file"; inp.accept = accept; inp.multiple = multiple; inp.style.display="none"; document.body.appendChild(inp); inp.addEventListener("change", ()=>{ cb([...inp.files]); inp.remove(); }); inp.click(); }
async function resizeImage(file, max, q=0.8){
  const url = URL.createObjectURL(file);
  try { const img = new Image(); await new Promise((res,rej)=>{ img.onload=res; img.onerror=rej; img.src=url; });
    const s = Math.min(1, max/Math.max(img.naturalWidth, img.naturalHeight)); const c = document.createElement("canvas");
    c.width = Math.round(img.naturalWidth*s); c.height = Math.round(img.naturalHeight*s); const g = c.getContext("2d"); g.fillStyle="#fff"; g.fillRect(0,0,c.width,c.height); g.drawImage(img,0,0,c.width,c.height);
    return c.toDataURL("image/jpeg", q); } finally { URL.revokeObjectURL(url); }
}
async function resizeLogo(file){
  const url = URL.createObjectURL(file);
  try { const img = new Image(); await new Promise((res,rej)=>{ img.onload=res; img.onerror=rej; img.src=url; });
    let w = img.naturalWidth || 600, h = img.naturalHeight || 200; const s = Math.min(1, 900/w, 300/h); w = Math.max(1,Math.round(w*s)); h = Math.max(1,Math.round(h*s));
    const c = document.createElement("canvas"); c.width = w; c.height = h; c.getContext("2d").drawImage(img, 0, 0, w, h);
    return c.toDataURL("image/png"); } finally { URL.revokeObjectURL(url); }
}
const readDataUrl = file => new Promise((res,rej)=>{ const r = new FileReader(); r.onload=()=>res(r.result); r.onerror=rej; r.readAsDataURL(file); });
async function addPhotos(id, files){
  const it = itemById(id); it.photos ||= []; let n = 0;
  for (const f of files.filter(f=>f.type.startsWith("image/"))){
    try { const data = await resizeImage(f, 1280, 0.78); it.photos.push({id:uid("p"),name:f.name,data,at:toLocalInput(now()),by:myName(),primary:!it.photos.length}); n++; }
    catch(e){ toast(`Couldn't read ${f.name}. Use a JPG or PNG photo.`); }
  }
  if (n){ itemLog(it, `${n} photo${n>1?"s":""} added`); save(); rerender(); toast(`${n} photo${n>1?"s":""} added`); }
}
async function addFiles(id, files){
  const it = itemById(id); it.files ||= []; let n = 0;
  for (const f of files){
    if (f.size > 3*1048576){ toast(`${f.name} is over 3 MB. Store it in SharePoint and paste the link in a comment.`); continue; }
    it.files.push({id:uid("f"),name:f.name,type:f.type||"application/octet-stream",size:f.size,data:await readDataUrl(f),at:toLocalInput(now()),by:myName()}); n++;
  }
  if (n){ save(); rerender(); toast(`${n} file${n>1?"s":""} attached`); }
}
function viewPhoto(id, i){
  const it = itemById(id), p = it.photos[i]; if (!p) return;
  openModal(`<header><h2>#${esc(id)} photo ${i+1} of ${it.photos.length}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody" style="align-items:center"><img src="${p.data}" alt="${esc(p.name||"")}" style="max-width:100%;max-height:70vh;border-radius:4px"><div class="small muted">${esc(p.name||"")} · ${fmtDate(p.at)}${p.by?" · "+esc(p.by):""}</div></div>
  <footer><button class="btn" data-act="a-photo-view" data-id="${esc(id)}" data-i="${(i-1+it.photos.length)%it.photos.length}" ${it.photos.length>1?"":"disabled"}>‹ Previous</button><button class="btn" data-act="a-photo-view" data-id="${esc(id)}" data-i="${(i+1)%it.photos.length}" ${it.photos.length>1?"":"disabled"}>Next ›</button></footer>`);
}
function openFile(id, i){
  const f = itemById(id).files[i];
  if (f.type.startsWith("image/")) { openModal(`<header><h2>${esc(f.name)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header><div class="mbody" style="align-items:center"><img src="${f.data}" alt="" style="max-width:100%;max-height:70vh"></div><footer><button class="btn" data-act="a-file-dl" data-id="${esc(id)}" data-i="${i}">Download</button></footer>`); return; }
  downloadData(f.name, f.data);
}
function downloadData(name, dataUrl){ saveFile(name, dataUrlToBlob(dataUrl)); }
document.addEventListener("dragover", e => { if (e.target.closest?.("#fp-wrap")) { e.preventDefault(); return; } const z = e.target.closest?.("[data-drop]"); if (z){ e.preventDefault(); z.classList.add("over"); } });
document.addEventListener("dragleave", e => { const z = e.target.closest?.("[data-drop]"); if (z && !z.contains(e.relatedTarget)) z.classList.remove("over"); });
document.addEventListener("drop", async e => { const fp = e.target.closest?.("#fp-wrap"); if (fp){ e.preventDefault(); const f = [...(e.dataTransfer?.files||[])].find(x=>x.type.startsWith("image/")); if (f){ try { pendingPhoto = await resizeImage(f, 900, 0.8); fp.innerHTML = photoPreview(pendingPhoto); } catch(_){ toast("Couldn't read that image. Use a JPG or PNG."); } } return; }
  const z = e.target.closest?.("[data-drop]"); if (!z) return; e.preventDefault(); z.classList.remove("over");
  const files = [...(e.dataTransfer?.files||[])]; if (!files.length) return;
  const need = {photos:"assets", mfiles:"shop"}[z.dataset.drop] || "notes";
  if (!can(need)) { toast("Your role can't add these"); return; }
  if (z.dataset.drop==="mfiles") { addDocs(files); return; }
  z.dataset.drop==="photos" ? addPhotos(z.dataset.id, files) : addFiles(z.dataset.id, files); });

/* ---- Maintenance / customer / location forms ---- */
let mdraft = null;
function docRow(f, i, editable, mid){
  const isImg = f.type.startsWith("image/");
  return `<div class="file">${isImg?`<img src="${f.data}" alt="" class="file-ic">`:`<span class="file-ic ext">${esc((f.name.split(".").pop()||"file").slice(0,4).toUpperCase())}</span>`}<div style="min-width:0;flex:1"><button type="button" class="linkbtn trunc" style="max-width:100%" data-act="m-file-open" data-i="${i}" ${mid?`data-mid="${mid}"`:""}>${esc(f.name)}</button><div class="small muted">${fmtBytes(f.size)} · ${fmtDate(f.at)}${f.by?" · "+esc(f.by):""}</div></div>${editable?`<button type="button" class="del" data-act="m-file-del" data-i="${i}" aria-label="Remove ${esc(f.name)}">×</button>`:""}</div>`;
}
function docsEditor(){
  const fl = mdraft.files;
  return `<div class="files">${fl.map((f,i)=>docRow(f,i,true)).join("")}</div>${fl.length?"":`<div class="ph-add file-add" style="min-height:52px;cursor:default"><span class="small muted">Inspection sheets, invoices, parts receipts, photos of damage or repairs. Up to 3 MB each. Drop files here or use Attach files.</span></div>`}`;
}
async function addDocs(files){
  let n = 0;
  for (const f of files){
    if (f.size > 3*1048576 && !f.type.startsWith("image/")){ toast(`${f.name} is over 3 MB. Store it in SharePoint and paste the link in the work notes.`); continue; }
    let data, type = f.type || "application/octet-stream", size = f.size, name = f.name;
    try {
      if (f.type.startsWith("image/") && f.type !== "image/gif"){ data = await resizeImage(f, 1600, 0.8); type = "image/jpeg"; size = Math.round(data.length*0.75); if (!/\.jpe?g$/i.test(name)) name = name.replace(/\.[^.]+$/,"") + ".jpg"; }
      else data = await readDataUrl(f);
    } catch(e){ toast(`Couldn't read ${f.name}.`); continue; }
    if (size > 3*1048576){ toast(`${f.name} is still over 3 MB after shrinking.`); continue; }
    mdraft.files.push({id:uid("d"),name,type,size,data,at:toLocalInput(now()),by:myName()}); n++;
  }
  if (n && mdraft.live){ const m = srById(mdraft.live); m.files = mdraft.files; srLog(m, `${n} document${n>1?"s":""} attached`); mdraft = null; save(); drawSR(); rerender(); toast(`${n} file${n>1?"s":""} attached`); return; }
  if (n && $("#m-files")){ $("#m-files").innerHTML = docsEditor(); toast(`${n} file${n>1?"s":""} attached. Save the work order to keep ${n>1?"them":"it"}.`); }
}
function viewDocs(mid){
  const m = S.maintenance.find(x=>x.id===mid), it = itemById(m.itemId);
  openModal(`<header><h2>Documents · #${esc(m.itemId)} ${esc(it?.name||"")}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><div class="small muted">${m.severity?m.severity+" severity · ":""}${esc(m.type)} work order · ${fmtDate(m.date+"T00:00")}${m.desc?" · "+esc(m.desc):""}</div><div class="files">${(m.files||[]).map((f,i)=>docRow(f,i,false,mid)).join("")}</div></div>
  <footer>${can("shop")?`<button class="btn" data-act="edit-maint" data-id="${mid}">Open work order</button>`:""}<button class="btn primary" data-act="close">Done</button></footer>`);
}
function docsBadge(m){ const n = (m.files||[]).length; return n ? `<button class="doc-badge" data-act="m-docs" data-id="${m.id}" title="${n} document${n>1?"s":""}" aria-label="${n} document${n>1?"s":""}">${ico("clip")}${n}</button>` : ""; }
function partsEditor(){
  if (!mdraft.parts.length) return `<div class="muted small">${S.inventory.length?"No parts used.":"Add parts on the Inventory page to use them here."}</div>`;
  return `<div style="display:flex;flex-direction:column;gap:8px">${mdraft.parts.map((p,i)=>{ const inv = S.inventory.find(x=>x.id===p.invId);
    return `<div class="charge-row" style="grid-template-columns:minmax(0,1fr) 70px 80px 28px"><select id="mp-i-${i}" data-mp="${i}" data-f="invId" aria-label="Part">${S.inventory.filter(x=>!x.retired||x.id===p.invId).map(x=>`<option value="${x.id}" ${x.id===p.invId?"selected":""}>${esc(x.name)} (${x.qty} ${esc(x.unit)})</option>`).join("")}</select>
    <input id="mp-q-${i}" type="number" step="0.1" min="0" data-mp="${i}" data-f="qty" value="${p.qty}" aria-label="Quantity"><span class="num small" style="text-align:right">${money((+p.qty||0)*(+p.cost||0))}</span><button class="del" data-act="m-delpart" data-i="${i}" aria-label="Remove">×</button></div>`; }).join("")}</div>`;
}
const POSITIONS = ["Assistant Superintendent","Superintendent","Senior Superintendent","Warranty Superintendent","Field Engineer","Project Engineer","Senior Field Engineer","Assistant Project Manager","Project Manager","Senior Project Manager"];
function simpleForm(kind, id, parent){
  const isC = kind==="cust"; const rec = id ? (isC?custById(id):locById(id)) : (isC?{name:"",company:"Bailey Harris",phone:"",email:"",notes:""}:{code:"",name:"",address:"",contact:"",notes:"",parent:parent||""});
  const fields = isC ? [["name","Name"],["position","Position"],["company","Company"],["phone","Phone"],["email","Email"],["notes","Notes"]] : [["code","Job #"],["name","Job name"],["parent","Part of"],["address","Address"],["contact","Site contact"],["notes","Notes"]];
  openModal(`<header><h2>${id?"Edit":"Add"} ${isC?"customer":"location"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">${isC&&rec.archived?`<div class="infobox">Archived ${rec.archivedAt?fmtDate(rec.archivedAt):""}${rec.archivedBy?" by "+esc(rec.archivedBy):""}. Hidden from new orders and check-outs.</div>`:""}<form id="simpleform" class="form" data-kind="${kind}" data-id="${id||""}">${fields.map(([k,l])=>k==="parent" && !isC ? (()=>{
      const kids = id ? locKids(id) : [], mine = id ? locFamily(id) : [], opts = sortedLocs().filter(x=>!mine.includes(x.id));
      return `<div class="field full"><label for="sf-parent">Part of</label><select id="sf-parent" name="parent"><option value="">None (top-level location)</option>${opts.map(x=>`<option value="${x.id}" ${rec.parent===x.id?"selected":""}>${esc(locOpt(x))}</option>`).join("")}</select><span class="hint">${kids.length?`Its ${kids.length} sub-location${kids.length>1?"s":""} (${kids.map(x=>esc(x.name)).join(", ")}) move with it.`:"Put a building under its job, a floor or room under its building, or a rack or pallet spot under its yard. Sub-locations can go as many levels deep as you need."}</span></div>`; })() : k==="contact" && !isC ? (()=>{
      const match = rec.contactId ? custById(rec.contactId) : S.customers.find(c=>!c.deleted && c.name===rec.contact);
      const legacy = !match && rec.contact ? rec.contact : "";
      const opts = S.customers.filter(c=>(!c.archived && !c.deleted) || c===match).slice().sort((a,b)=>a.name.localeCompare(b.name));
      return `<div class="field"><label for="sf-contact">Site contact</label><select id="sf-contact" name="contact"><option value="">No site contact</option>${opts.map(c=>`<option value="${c.id}" ${match===c?"selected":""}>${esc(c.name)}${c.position?` · ${esc(c.position)}`:""}${c.archived?" (archived)":""}</option>`).join("")}${legacy?`<option value="__keep" selected>${esc(legacy)} (not in Customers)</option>`:""}</select><span class="hint">Pick from your Customers. Add someone on the Customers page first if they're missing.</span></div>`; })() : k==="position" ? `<div class="field"><label for="sf-position">Position</label><select id="sf-position" name="position"><option value="">Select a position</option>${POSITIONS.concat(rec.position && !POSITIONS.includes(rec.position) ? [rec.position] : []).map(p=>`<option ${rec.position===p?"selected":""}>${esc(p)}</option>`).join("")}</select></div>` : `<div class="field ${k==="notes"||k==="address"?"full":""}"><label for="sf-${k}">${l}</label><input id="sf-${k}" name="${k}" value="${esc(rec[k])}"></div>`).join("")}</form></div>
  <footer>${isC && id && me()?.role==="Admin" ? `<span class="foot-left"><button class="btn danger" data-act="cust-del" data-id="${id}">Delete</button>${rec.archived?`<button class="btn" data-act="cust-restore" data-id="${id}">Restore</button>`:`<button class="btn" data-act="cust-archive" data-id="${id}">Archive</button>`}</span>` : ""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="save-simple">Save</button></footer>`,"narrow");
}
function saveSimple(){ const f=$("#simpleform"); const fd=Object.fromEntries(new FormData(f)); const isC=f.dataset.kind==="cust"; const arr=isC?S.customers:S.locations;
  if (!isC){ if ("parent" in fd){ if (!fd.parent || fd.parent===f.dataset.id || (f.dataset.id && locFamily(f.dataset.id).includes(fd.parent))) { fd.parent = ""; } } }
  if (!isC){ if (fd.contact==="__keep"){ delete fd.contact; } else { const c = custById(fd.contact); fd.contactId = c ? c.id : ""; fd.contact = c ? c.name : ""; } }
  if (!(fd.name||"").trim()){ toast("Name is required"); return; }
  if (f.dataset.id){ const r = arr.find(x=>x.id===f.dataset.id); Object.assign(r, fd); if (!isC && !r.parent) delete r.parent; }
  else { const r = {id:uid(isC?"c-":"l-"),...fd}; if (!isC && !r.parent) delete r.parent; arr.push(r); if (!isC && r.parent && route()==="locations") ui.locs.open = r.parent; }
  save(); closeModal(); toast("Saved"); rerender(); }

/* ================= Check In / Out (internal custody, no billing) ================= */
ui.co = {tab:"out", q:"", sel:new Set()};
function holderNames(){ return [...new Set(S.customers.filter(c=>!c.archived&&!c.deleted).map(c=>c.name).concat(S.custody.map(c=>c.holder)).filter(Boolean))].sort(); }
function vCheckout(){
  const u = ui.co, open = S.custody.filter(c=>!c.inAt), over = open.filter(custodyOverdue), hist = S.custody.filter(c=>c.inAt);
  let list = u.tab==="out" ? open : u.tab==="overdue" ? over : hist;
  if (u.q){ const q = u.q.toLowerCase(); list = list.filter(c=>[c.ref,custodyLabel(c),c.holder,locLabel(locById(c.locationId)),c.outNotes].join(" ").toLowerCase().includes(q)); }
  list = list.slice().sort((a,b)=> u.tab==="hist" ? b.inAt.localeCompare(a.inAt) : (a.dueAt||"9999").localeCompare(b.dueAt||"9999"));
  const byHolder = {}; open.forEach(c=>{ const h = byHolder[c.holder] ||= {n:0,late:0}; h.n += c.kind==="stock"?1:1; if (custodyOverdue(c)) h.late++; });
  const holders = Object.entries(byHolder).sort((a,b)=>b[1].n-a[1].n);
  const allSel = list.length && list.every(c=>u.sel.has(c.id));
  return band("Check In / Out","Assets, stock and inventory your own crews take out. Nothing here is billed. Checked-out assets can't be rented until they come back, and inventory comes off the shelf until it's returned.",`<button class="btn" data-act="export-co">Export to CSV</button><button class="btn" data-act="co-new" data-in="1" ${open.length?"":"disabled"}>Check in</button><button class="btn primary" data-act="co-new">+ Check out</button>`) +
  tabs([["out","Checked out",open.length],["overdue","Overdue",over.length],["hist","Returned",hist.length]],u.tab,"cotab") + `
  <div class="content">
    <div class="kpis">
      <div class="kpi"><span class="lbl">Out now</span><span class="val num">${open.length}</span><span class="note">${open.filter(c=>c.kind==="asset").length} assets · ${open.filter(c=>c.kind==="stock").length} stock · ${open.filter(c=>c.kind==="inv").length} inventory</span></div>
      <div class="kpi ${over.length?"bad":""}"><span class="lbl">Past due</span><span class="val num">${over.length}</span><span class="note">past the due-back date</span></div>
      <div class="kpi"><span class="lbl">People holding gear</span><span class="val num">${holders.length}</span><span class="note">${holders.slice(0,2).map(h=>esc(h[0])).join(", ")||"nobody right now"}</span></div>
    </div>
    <div class="toolbar">
      ${u.tab!=="hist"?`<button class="btn primary" data-act="co-bulkin" ${u.sel.size?"":"disabled"}>Check in selected (${u.sel.size})</button>`:""}
      <div class="search">${ico("search")}<input id="coq" type="search" placeholder="Search person, asset, job or ref" value="${esc(u.q)}" data-input="coq"></div>
    </div>
    <div class="tablewrap"><table>
      <thead><tr>${u.tab!=="hist"?`<th style="width:36px"><input type="checkbox" data-act="co-selall" ${allSel?"checked":""} aria-label="Select all"></th>`:""}<th>Ref</th><th>Item</th><th style="text-align:right">Qty</th><th>Assigned to</th><th>Job / location</th><th>Checked out</th><th>${u.tab==="hist"?"Returned":"Due back"}</th><th>Status</th><th>Notes</th><th></th></tr></thead>
      <tbody>${list.length?list.map(c=>{ const st = c.inAt?"Returned":custodyOverdue(c)?"Overdue":"Checked Out";
        return `<tr>${u.tab!=="hist"?`<td><input type="checkbox" data-act="co-sel" data-id="${c.id}" ${u.sel.has(c.id)?"checked":""} aria-label="Select ${esc(c.ref)}"></td>`:""}
        <td class="num"><b>${esc(c.ref)}</b></td>
        <td>${c.kind==="asset"?`<button class="linkbtn" data-act="view-item" data-id="${esc(c.itemId)}">#${esc(c.itemId)}</button> ${esc(itemById(c.itemId)?.name||"")}`:`${esc(custodyLabel(c))} <span class="chip ${c.kind==="inv"?"maint":"checked"}">${c.kind==="inv"?"inventory":"stock"}</span>`}</td>
        <td class="num" style="text-align:right">${c.qty||1}${c.kind==="inv"?" "+esc(S.inventory.find(x=>x.id===c.invId)?.unit||""):""}${!c.inAt&&c.backSoFar?`<div class="small muted">${c.backSoFar} already back</div>`:""}${c.inAt&&c.kind!=="asset"&&c.returnedQty!=null&&c.returnedQty!==c.qty?`<div class="small" style="color:${c.kind==="inv"?"var(--muted)":"var(--bad)"}">${c.returnedQty} back · ${c.kind==="inv"?(c.usedQty+" used"):(c.missingQty+" missing")}</div>`:""}</td>
        <td><b>${esc(c.holder)}</b></td><td>${esc(locLabel(locById(c.locationId)))}</td>
        <td class="num">${fmtDT(c.outAt)}</td><td class="num">${c.inAt?fmtDT(c.inAt):c.dueAt?fmtDT(c.dueAt):'<span class="muted">Open-ended</span>'}</td>
        <td>${chipFor(st)}</td><td><span class="trunc" title="${esc((c.outNotes||"")+(c.inNotes?" / In: "+c.inNotes:""))}">${esc(c.inNotes||c.outNotes||"")}</span></td>
        <td>${c.inAt?`<button class="btn sm" data-act="co-undo" data-id="${c.id}">Undo check-in</button>`:`<button class="btn sm" data-act="co-in" data-id="${c.id}">Check in</button>`}</td></tr>`; }).join(""):`<tr><td colspan="11"><div class="empty"><strong>${u.tab==="hist"?"Nothing has been returned yet":u.tab==="overdue"?"Nothing is past due":"Nothing is checked out"}</strong>Use <b>Check out</b> to hand assets, stock or inventory to a crew member. Pick the type from the drop-down and add as many items as you need.</div></td></tr>`}</tbody>
    </table></div>
    ${holders.length?`<div class="panel"><h3>Who has what</h3><div class="body"><div class="bars">${holders.map(([h,v])=>`<div class="bar-row"><span>${esc(h)}</span><div class="track"><div class="fill" style="width:${v.n/holders[0][1].n*100}%;${v.late?"background:var(--bad)":""}"></div></div><span class="v">${v.n} out${v.late?` · ${v.late} late`:""}</span></div>`).join("")}</div></div></div>`:""}
  </div>`;
}
let codraft = null;
const CO_TYPES = [["asset","Assets"],["stock","Asset stock"],["inv","Inventory"]];
function invById(id){ return S.inventory.find(x=>x.id===id); }
function invOut(invId){ return S.custody.filter(c=>c.kind==="inv" && c.invId===invId && !c.inAt).reduce((s,c)=>s+(+c.qty||0),0); }
function openCheckoutForm(pre={}){
  const d = now(); const due = new Date(d.getTime()+7*DAY); due.setHours(17,0,0,0);
  codraft = {holder:"",locationId:yardId(),outAt:toLocalInput(d),dueAt:toLocalInput(due),notes:"",
    assets:pre.assets||[], stock:pre.stock||[], inv:pre.inv||[], pq:"", type:pre.type || (pre.inv?.length?"inv":pre.stock?.length?"stock":"asset")};
  drawCheckoutForm();
}
function coLineCount(c){ return c.assets.length + c.stock.filter(s=>+s.qty>0).length + c.inv.filter(s=>+s.qty>0).length; }
function coValid(c){
  if (!c.holder.trim() || !coLineCount(c) || (c.dueAt && parseD(c.dueAt) <= parseD(c.outAt))) return false;
  if (c.stock.some(s=>{ const st=stockById(s.stockId); return !st || +s.qty > st.qty - stockOut(st.id) || +s.qty < 0; })) return false;
  if (c.inv.some(s=>{ const it=invById(s.invId); return !it || +s.qty > it.qty || +s.qty < 0; })) return false;
  return true;
}
function coPicker(c){
  const q = c.pq.toLowerCase();
  if (c.type==="asset"){
    const rows = S.items.filter(it=>it.status!=="Retired").filter(it=>!q||(it.id+" "+it.name+" "+it.category).toLowerCase().includes(q)).map(it=>{
      const sel = c.assets.includes(it.id), st = itemStatus(it), busy = !sel && st!=="Available";
      const res = nextReservation(it.id); const warn = !busy && res && c.dueAt && parseD(res.out) < parseD(c.dueAt);
      return `<label class="pick ${busy?"busy":""}"><input type="checkbox" data-act="co-pick" data-id="${esc(it.id)}" ${sel?"checked":""} ${busy?"disabled":""}><span class="num muted">#${esc(it.id)}</span><span class="trunc" style="max-width:none">${esc(it.name)}</span><span class="small ${warn?"":"muted"}" style="${warn?"color:var(--warn)":""}">${busy?esc(st):warn?`Booked ${fmtDate(res.out)}`:""}</span></label>`; }).join("");
    return rows || '<div class="empty">No assets match.</div>';
  }
  const isInv = c.type==="inv", src = isInv ? S.inventory.filter(x=>!x.retired) : S.stock, lines = isInv ? c.inv : c.stock, key = isInv ? "invId" : "stockId";
  if (!src.length) return `<div class="empty"><strong>No ${isInv?"inventory":"stock"} items yet</strong>Add them on the ${isInv?"Inventory":"Asset Stock"} page first.</div>`;
  const rows = src.filter(x=>!q||(x.name+" "+(x.sku||"")+" "+(x.category||"")).toLowerCase().includes(q)).map(x=>{
    const avail = isInv ? +x.qty : x.qty - stockOut(x.id), inCart = lines.find(l=>l[key]===x.id), none = avail <= 0;
    return `<div class="pick qty-pick ${none&&!inCart?"busy":""}">${photoThumb(x.photo, isInv?"inv":"stock", x.id, isInv?"inventory":"stock").replace('style="width:40px;height:40px"','style="width:30px;height:30px"')}
      <span class="trunc" style="max-width:none"><b>${esc(x.name)}</b>${x.sku?` <span class="muted small">${esc(x.sku)}</span>`:""}</span>
      <span class="small num ${none?"":"muted"}" style="${none?"color:var(--bad)":""}">${avail} ${esc(x.unit||"")} ${isInv?"on hand":"available"}</span>
      ${inCart?`<span class="chip available">Added ×${inCart.qty}</span>`:`<span class="qty-add"><input id="qa-${x.id}" type="number" min="1" ${isInv?'step="any"':'step="1"'} max="${avail}" value="1" aria-label="Quantity of ${esc(x.name)}" ${none?"disabled":""}><button class="btn sm" data-act="co-add" data-kind="${c.type}" data-id="${x.id}" ${none?"disabled":""}>Add</button></span>`}</div>`; }).join("");
  return rows || '<div class="empty">Nothing matches.</div>';
}
function coCart(c){
  const rows = [];
  c.assets.forEach(id=>{ const it = itemById(id); rows.push(`<div class="line cart-line"><div><span class="chip neutral">Asset</span> <b>#${esc(id)}</b> ${esc(it?.name||"")}</div><span class="num muted small">1</span><button class="del" data-act="co-line-del" data-kind="asset" data-id="${esc(id)}" aria-label="Remove">×</button></div>`); });
  [["stock",c.stock,"stockId",stockById,"Stock"],["inv",c.inv,"invId",invById,"Inventory"]].forEach(([k,list,key,get,label])=>list.forEach((l,i)=>{ const x = get(l[key]); const avail = !x ? 0 : k==="inv" ? +x.qty : x.qty - stockOut(x.id); const over = +l.qty > avail;
    rows.push(`<div class="line cart-line"><div><span class="chip ${k==="inv"?"maint":"checked"}">${label}</span> <b>${esc(x?.name||"(deleted)")}</b>${over?`<div class="small" style="color:var(--bad)">Only ${avail} ${esc(x?.unit||"")} ${k==="inv"?"on hand":"available"}</div>`:""}</div><span class="cart-qty"><input id="cq-${k}-${i}" type="number" min="1" step="${k==="inv"?"any":"1"}" data-cq="${k}" data-i="${i}" value="${l.qty}" aria-label="Quantity"> <span class="small muted">${esc(x?.unit||"")}</span></span><button class="del" data-act="co-line-del" data-kind="${k}" data-i="${i}" aria-label="Remove">×</button></div>`); }));
  return rows.length ? `<div class="lines">${rows.join("")}</div>` : '<div class="muted small">Nothing added yet. Pick a type above, then add items.</div>';
}
function drawCheckoutForm(){
  const c = codraft;
  const counts = {asset:c.assets.length, stock:c.stock.length, inv:c.inv.length};
  openModal(`<header><h2>Check out</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">
    <div class="form">
      <div class="field"><label for="co-holder">Assigned to</label><input id="co-holder" list="holder-list" data-co="holder" value="${esc(c.holder)}" placeholder="Name or crew"><datalist id="holder-list">${holderNames().map(h=>`<option value="${esc(h)}">`).join("")}</datalist></div>
      <div class="field"><label for="co-loc">Job / location</label><select id="co-loc" data-co="locationId">${sortedLocs().map(l=>`<option value="${l.id}" ${l.id===c.locationId?"selected":""}>${esc(locOpt(l))}</option>`).join("")}</select></div>
      <div class="field"><label for="co-out">Checked out</label><input id="co-out" type="datetime-local" data-co="outAt" value="${esc(c.outAt)}"></div>
      <div class="field"><label for="co-due">Due back</label><input id="co-due" type="datetime-local" data-co="dueAt" value="${esc(c.dueAt)}"><span class="hint">Clear it for open-ended</span></div>
      <div class="field" style="grid-column:span 2"><label for="co-notes">Notes</label><input id="co-notes" data-co="notes" value="${esc(c.notes)}" placeholder="Condition, hours, accessories sent"></div>
    </div>
    <div class="subhead"><span>Add items</span></div>
    <div class="co-pickbar">
      <div class="field"><label for="co-type">Item type</label><select id="co-type" data-co="type">${CO_TYPES.map(([k,l])=>`<option value="${k}" ${c.type===k?"selected":""}>${l}${counts[k]?` (${counts[k]} added)`:""}</option>`).join("")}</select></div>
      <div class="field" style="flex:1"><label for="co-pq">Search</label><input id="co-pq" type="search" data-co="pq" value="${esc(c.pq)}" placeholder="${c.type==="asset"?"Asset number or name":c.type==="stock"?"Stock item name":"Part, fuel or SKU"}"></div>
    </div>
    <div class="hint" style="margin-top:-8px">${c.type==="asset"?"Only assets in the yard can go out.":c.type==="stock"?"Stock comes back by count. Anything not returned is recorded as missing.":"Inventory is taken off the shelf now. On check-in, enter what came back; the rest is logged as used."}</div>
    <div class="picker">${coPicker(c)}</div>
    <div class="subhead"><span>Checking out (${coLineCount(c)})</span></div>
    <div id="co-cart">${coCart(c)}</div>
  </div>
  <footer><span class="total">${coLineCount(c)} line${coLineCount(c)===1?"":"s"}</span><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="co-save" ${coValid(c)?"":"disabled"}>Check out</button></footer>`);
}
function saveCheckout(){
  const c = codraft;
  for (const s of c.stock){ const st = stockById(s.stockId); if (+s.qty > st.qty - stockOut(st.id)){ toast(`Only ${st.qty - stockOut(st.id)} ${st.name} available`); return; } }
  for (const s of c.inv){ const it = invById(s.invId); if (+s.qty > it.qty){ toast(`Only ${it.qty} ${it.unit} of ${it.name} on hand`); return; } }
  const ref = "CO-" + (S.nextCheckout++), at = toLocalInput(now());
  const base = {ref,holder:c.holder.trim(),locationId:c.locationId,outAt:c.outAt,dueAt:c.dueAt||null,outNotes:c.notes,inAt:null,inNotes:"",by:myName()};
  c.assets.forEach(id=>S.custody.push({id:uid("k"),kind:"asset",itemId:id,qty:1,...base}));
  c.stock.filter(s=>+s.qty>0).forEach(s=>S.custody.push({id:uid("k"),kind:"stock",stockId:s.stockId,qty:+s.qty,...base}));
  c.inv.filter(s=>+s.qty>0).forEach(s=>{ const it = invById(s.invId); const q = +s.qty;
    it.qty = +(it.qty - q).toFixed(2); (it.log ||= []).push({by:myName(),at,delta:-q,reason:`Checked out to ${base.holder} (${ref})`});
    S.custody.push({id:uid("k"),kind:"inv",invId:s.invId,qty:q,...base}); });
  codraft = null; save(); closeModal(); toast(`${ref} checked out to ${base.holder}`); rerender();
}
let coin = null;
function openCustodyCheckin(ids){
  ids = ids.filter(id=>S.custody.some(c=>c.id===id && !c.inAt));
  if (!ids.length){ toast("Pick at least one checked-out line"); return; }
  const back = {}, rest = {};
  ids.forEach(id=>{ const c = S.custody.find(x=>x.id===id); if (c.kind!=="asset"){ back[id] = c.qty; rest[id] = "out"; } });
  const loc = {}; ids.forEach(id=>{ const c = S.custody.find(x=>x.id===id); if (c.kind==="asset" && itemById(c.itemId)) loc[id] = itemById(c.itemId).locationId || yardId(); });
  coin = {ids, at:toLocalInput(now()), notes:"", maint:{}, back, rest, loc}; drawCustodyCheckin();
}
function coinUnit(c){ return c.kind==="inv" ? invById(c.invId)?.unit||"" : c.kind==="stock" ? stockById(c.stockId)?.unit||"" : ""; }
function coinRestHtml(c){
  const left = +(c.qty - (+coin.back[c.id]||0)).toFixed(2), unit = esc(coinUnit(c));
  if (!(left > 0)) return "";
  return `<div class="rest-row"><span class="small"><b class="num">${left} ${unit}</b> not returned:</span>
    <select id="cr-${c.id}" data-coin-rest="${c.id}" aria-label="What happened to the rest">
      <option value="out" ${coin.rest[c.id]==="out"?"selected":""}>Still checked out to ${esc(c.holder)}</option>
      <option value="gone" ${coin.rest[c.id]==="gone"?"selected":""}>${c.kind==="inv"?"Used up on the job":"Missing or lost"}</option>
    </select></div>`;
}
function drawCustodyCheckin(){
  const recs = coin.ids.map(id=>S.custody.find(c=>c.id===id));
  const row = c => {
    const unit = esc(coinUnit(c));
    const right = c.kind==="asset" ? `<span class="ret-col"><label class="small muted" for="rl-${c.id}">Returning to</label>${returnLocSelect("data-coin-loc", c.id, coin.loc[c.id]).replace('<select ','<select id="rl-'+c.id+'" ')}<label class="small" style="display:flex;gap:6px;align-items:center"><input type="checkbox" data-act="coin-maint" data-id="${c.id}" ${coin.maint[c.id]?"checked":""} style="accent-color:var(--accent)"> Needs service</label></span>`
      : `<span class="cart-qty"><label class="small muted" for="cb-${c.id}">Returning</label><input id="cb-${c.id}" type="number" min="0" max="${c.qty}" step="${c.kind==="inv"?"any":"1"}" data-coin-q="${c.id}" value="${coin.back[c.id]}"> <span class="small muted">of ${c.qty} ${unit}</span></span>`;
    return `<div class="line coin-line"><div><span class="chip ${c.kind==="inv"?"maint":c.kind==="stock"?"checked":"neutral"}">${c.kind==="inv"?"Inventory":c.kind==="stock"?"Stock":"Asset"}</span> <b>${esc(custodyLabel(c))}</b><div class="calc">${esc(c.ref)} · ${esc(c.holder)} · out ${fmtDate(c.outAt)}${c.backSoFar?` · ${c.backSoFar} ${unit} already returned`:""}${custodyOverdue(c)?' · <span style="color:var(--bad)">overdue</span>':""}</div>${c.kind!=="asset"?`<div id="cr-wrap-${c.id}">${coinRestHtml(c)}</div>`:""}</div><span></span>${right}</div>`; };
  openModal(`<header><h2>Check in ${recs.length} line${recs.length>1?"s":""}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">
    <div class="lines">${recs.map(row).join("")}</div>
    <div class="form"><div class="field"><label for="coin-at">Returned at</label><input id="coin-at" type="datetime-local" data-coin="at" value="${esc(coin.at)}"></div>
    <div class="field" style="grid-column:span 2"><label for="coin-notes">Condition notes</label><input id="coin-notes" data-coin="notes" value="${esc(coin.notes)}" placeholder="Damage, missing parts, cleaning needed"></div></div>
    <div class="small muted">For stock and inventory, enter how many are coming back now. Anything not returned stays checked out unless you mark it ${"missing or used"}. Assets marked <b>Needs service</b> get an open work order under Maintenance.</div>
  </div>
  <footer><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="coin-save">Check in</button></footer>`);
}
function saveCustodyCheckin(){
  for (const id of coin.ids){ const c = S.custody.find(x=>x.id===id); if (c.kind!=="asset"){ const b = +coin.back[id]; if (coin.back[id]==="" || isNaN(b) || b < 0 || b > c.qty){ toast(`Returning amount for ${custodyLabel(c)} must be between 0 and ${c.qty}`); return; } } }
  const by = myName(), at = coin.at; let closed = 0, partial = 0, skipped = 0;
  coin.ids.forEach(id=>{
    const c = S.custody.find(x=>x.id===id);
    if (c.kind==="asset"){
      Object.assign(c, {inAt:at, inNotes:coin.notes, inBy:by}); closed++;
      if (coin.loc?.[id]) applyReturnLocation(c.itemId, coin.loc[id], `check-out ${c.ref}`);
      if (coin.maint[id]) newSR({itemId:c.itemId,date:at.slice(0,10),type:"Corrective maintenance",severity:"Medium",opStatus:"Degraded",stage:"Accepted",desc:`Returned from ${c.holder} (${c.ref})${coin.notes?": "+coin.notes:""}`,why:`crew check-in ${c.ref}`});
      return;
    }
    const b = +(+coin.back[id]).toFixed(2), left = +(c.qty - b).toFixed(2), keepOut = left > 0 && coin.rest[id] === "out";
    const unit = coinUnit(c);
    if (c.kind==="inv" && b){ const it = invById(c.invId); if (it){ it.qty = +(it.qty + b).toFixed(2); (it.log ||= []).push({by,at,delta:b,reason:`Returned by ${c.holder} (${c.ref})`}); } }
    if (keepOut){
      if (!b){ skipped++; return; }
      // split: a returned record for what came back, the original stays open with the rest
      S.custody.push({...c, id:uid("k"), qty:b, inAt:at, inNotes:coin.notes, inBy:by, returnedQty:b, partOf:c.id});
      c.qty = left; c.backSoFar = +((c.backSoFar||0) + b).toFixed(2); partial++;
      return;
    }
    Object.assign(c, {inAt:at, inNotes:coin.notes, inBy:by, returnedQty:b});
    if (left > 0){
      if (c.kind==="inv"){ c.usedQty = left; const it = invById(c.invId); if (it) (it.log ||= []).push({by,at,delta:0,reason:`${left} ${unit} used on the job by ${c.holder} (${c.ref})`}); }
      else { c.missingQty = left; const st = stockById(c.stockId); if (st){ st.qty = Math.max(0, st.qty - left); (st.log ||= []).push({by,at,delta:-left,reason:`Missing after ${c.ref} (${c.holder})`}); } }
    }
    closed++;
  });
  coin = null; ui.co.sel.clear(); save(); closeModal(); rerender();
  toast([closed?`${closed} line${closed>1?"s":""} checked in`:"", partial?`${partial} partly returned, the rest is still checked out`:"", skipped?`${skipped} unchanged (nothing returned)`:""].filter(Boolean).join(" · ") || "Nothing changed");
}
function undoCheckin(id){
  const c = S.custody.find(x=>x.id===id); if (!c || !c.inAt) return;
  const parent = c.partOf ? S.custody.find(x=>x.id===c.partOf && !x.inAt) : null;
  if (c.kind==="asset" && openCustody(c.itemId)){ toast(`#${c.itemId} is checked out again, so this can't be undone`); return; }
  if (c.kind==="asset" && activeOrderFor(c.itemId)){ toast(`#${c.itemId} is on a rental now, so this can't be undone`); return; }
  if (c.kind==="inv"){ const it = invById(c.invId), b = +c.returnedQty||0; if (it && b){ if (it.qty < b){ toast(`Only ${it.qty} ${it.unit} of ${it.name} on hand, so the return can't be reversed`); return; } it.qty = +(it.qty - b).toFixed(2); (it.log ||= []).push({by:myName(),at:toLocalInput(now()),delta:-b,reason:`Check-in undone (${c.ref})`}); } }
  if (c.kind==="stock" && c.missingQty){ const st = stockById(c.stockId); if (st){ st.qty += c.missingQty; (st.log ||= []).push({by:myName(),at:toLocalInput(now()),delta:c.missingQty,reason:`Missing count reversed (${c.ref})`}); } }
  if (parent){ parent.qty = +(parent.qty + c.qty).toFixed(2); parent.backSoFar = Math.max(0, +((parent.backSoFar||0) - c.qty).toFixed(2)) || undefined; S.custody = S.custody.filter(x=>x.id!==c.id); }
  else { const q = c.kind!=="asset" ? +(c.qty).toFixed(2) : c.qty; Object.assign(c, {inAt:null, inNotes:"", inBy:undefined, returnedQty:undefined, usedQty:undefined, missingQty:undefined, qty:q}); }
  save(); rerender(); toast(`Check-in undone. ${custodyLabel(c)} is checked out to ${c.holder} again.`);
}

/* ================= Asset Optimizer ================= */
function lastUsed(itemId){
  let last = null;
  S.orders.filter(o=>["Rented Out","Returned"].includes(o.status)&&o.lines.some(l=>l.itemId===itemId)).forEach(o=>{ const e = o.status==="Rented Out" ? toLocalInput(now()) : (o.returnedAt||o.ret); if (!last||e>last) last=e; });
  S.custody.filter(c=>c.kind==="asset"&&c.itemId===itemId).forEach(c=>{ const e = c.inAt||toLocalInput(now()); if (!last||e>last) last=e; });
  return last;
}
function custodyUse(itemId, days=90){ const end=now(), start=new Date(end-days*DAY); let used=0;
  S.custody.filter(c=>c.kind==="asset"&&c.itemId===itemId).forEach(c=>{ const s=parseD(c.outAt), e=c.inAt?parseD(c.inAt):end; const ov=Math.min(e,end)-Math.max(s,start); if (ov>0) used+=ov; });
  return Math.min(1, used/(days*DAY)); }
/* ---- One main photo for stock and inventory items ---- */
let pendingPhoto;  // undefined = unchanged, null = removed, string = new image
function photoThumb(photo, kind, id, icon){ return photo ? `<button class="thumb-btn" data-act="x-photo-view" data-kind="${kind}" data-id="${id}" aria-label="View photo"><img class="thumb" src="${photo}" alt="" style="width:40px;height:40px"></button>` : `<span class="thumb ph" style="width:40px;height:40px">${ico(icon,1.6)}</span>`; }
function photoField(current){ pendingPhoto = undefined; return `<div class="field full"><label>Photo</label><div class="photo-row" id="fp-wrap">${photoPreview(current)}</div></div>`; }
function photoPreview(p){ return `<div class="fp-img">${p?`<img src="${p}" alt="">`:ico("camera",1.6)}</div><div style="display:flex;flex-direction:column;gap:6px"><div style="display:flex;gap:6px;flex-wrap:wrap"><button type="button" class="btn sm" data-act="x-photo-pick">${p?"Change photo":"Upload photo"}</button>${p?`<button type="button" class="btn sm" data-act="x-photo-clear">Remove</button>`:""}</div><span class="hint">JPG or PNG, or drop one here. It's resized to save space.</span></div>`; }
function applyPendingPhoto(rec){ if (pendingPhoto !== undefined){ if (pendingPhoto) rec.photo = pendingPhoto; else delete rec.photo; } pendingPhoto = undefined; }
function viewXPhoto(kind, id){ const rec = kind==="inv" ? S.inventory.find(x=>x.id===id) : stockById(id); if (!rec?.photo) return;
  openModal(`<header><h2>${esc(rec.name)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header><div class="mbody" style="align-items:center"><img src="${rec.photo}" alt="${esc(rec.name)}" style="max-width:100%;max-height:70vh;border-radius:4px"></div><footer><button class="btn" data-act="${kind==="inv"?"inv-edit":"stock-edit"}" data-id="${id}">Edit item</button></footer>`); }

/* ================= Asset Stock (quantity-tracked) ================= */
function vStock(){
  const rows = S.stock.map(s=>({s,out:stockOut(s.id)}));
  return band("Asset Stock","Items you track by quantity instead of one by one, such as cones, barricades, cords or shoring. Check them out by the count.",`<button class="btn primary" data-act="stock-new">+ Add stock item</button>`) + `
  <div class="content">
    <div class="kpis"><div class="kpi"><span class="lbl">Items tracked</span><span class="val num">${rows.length}</span><span class="note">${rows.reduce((t,r)=>t+(+r.s.qty||0),0)} units owned</span></div>
    <div class="kpi"><span class="lbl">Checked out</span><span class="val num">${rows.reduce((t,r)=>t+r.out,0)}</span><span class="note">units out with crews</span></div>
    <div class="kpi ${rows.some(r=>(+r.s.qty||0)-r.out<=0)?"bad":""}"><span class="lbl">None available</span><span class="val num">${rows.filter(r=>(+r.s.qty||0)-r.out<=0).length}</span><span class="note">items with 0 left to check out</span></div></div>
    <div class="tablewrap" style="margin-top:14px"><table><thead><tr><th>Item</th><th>Group</th><th style="text-align:right">Total</th><th style="text-align:right">Checked out</th><th style="text-align:right">Available</th><th>Rates</th><th>Out with</th><th>Kept at</th><th>Notes</th><th></th></tr></thead><tbody>
  ${rows.length?rows.map(({s,out})=>`<tr><td><div class="idcell">${photoThumb(s.photo,"stock",s.id,"stock")}<b>${esc(s.name)}</b></div></td><td class="muted" style="white-space:nowrap" title="${esc(groupLabel(s.category||""))}">${esc(s.category)||"—"}</td><td class="num" style="text-align:right">${s.qty} ${esc(s.unit)}</td><td class="num" style="text-align:right">${out}</td><td class="num" style="text-align:right">${s.qty-out<=0?chipFor("Out of stock"):`<b>${s.qty-out}</b>`}</td><td class="small num">${rateParts(s).map(([l,v])=>`${l[0]} ${money0(v)}`).join(" · ")||'<span class="muted">—</span>'}</td><td class="small">${S.custody.filter(c=>c.kind==="stock"&&c.stockId===s.id&&!c.inAt).map(c=>`${esc(c.holder)} <span class="num muted">×${c.qty}</span>${custodyOverdue(c)?' <span style="color:var(--bad)">late</span>':""}`).join("<br>")||'<span class="muted">—</span>'}</td><td>${esc(keptAt(s))}</td><td><span class="trunc">${esc(s.notes)}</span></td><td><div style="display:flex;gap:6px;justify-content:flex-end"><button class="btn sm" data-act="stock-co" data-id="${s.id}" ${s.qty-out>0?"":"disabled"}>Check out</button><button class="btn sm" data-act="stock-edit" data-id="${s.id}">Edit</button></div></td></tr>`).join(""):`<tr><td colspan="9"><div class="empty"><strong>No stock items yet</strong>Add things you own in quantity, like 40 traffic cones or 12 extension cords. They can then be checked out by count from Check In/Out.</div></td></tr>`}
  </tbody></table></div></div>`;
}
function stockForm(id){
  const s = id ? stockById(id) : {name:"",category:(S.groups.find(g=>g.name==="Site Services")||S.groups[0])?.name||"Tools",qty:1,unit:"ea",locationId:yardId(),notes:""};
  openModal(`<header><h2>${id?"Edit stock item":"Add stock item"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="stockform" class="form" data-id="${id||""}">
    <div class="field full"><label for="sk-name">Name</label><input id="sk-name" name="name" value="${esc(s.name)}" placeholder="Traffic cones 28in"></div>
    <div class="field"><label for="sk-cat">Group</label><select id="sk-cat" name="category">${groupOptions(s.category)}</select></div>
    <div class="field"><label for="sk-qty">Total owned</label><input id="sk-qty" type="number" min="0" name="qty" value="${s.qty}"></div>
    <div class="field"><label for="sk-unit">Unit</label><input id="sk-unit" name="unit" value="${esc(s.unit)}"></div>
    <div class="field"><label for="sk-loc">Kept at</label>${keptAtSelect("sk-loc", s.locationId)}</div>
    <div class="field full"><label>Rental rates per ${esc(s.unit||"unit")} <span class="muted" style="text-transform:none;letter-spacing:0;font-weight:400">(optional, shown to superintendents)</span></label><div class="rate-trio"><label>Day<input type="number" min="0" step="0.01" name="day" value="${s.day??""}" placeholder="0.00"></label><label>Week<input type="number" min="0" step="0.01" name="week" value="${s.week??""}" placeholder="0.00"></label><label>Month<input type="number" min="0" step="0.01" name="month" value="${s.month??""}" placeholder="0.00"></label></div></div>
    <div class="field full"><label for="sk-notes">Notes</label><input id="sk-notes" name="notes" value="${esc(s.notes)}"></div>
    ${photoField(s.photo)}
  </form></div>
  <footer>${id?`<button class="btn danger" data-act="stock-del" data-id="${id}" style="margin-right:auto">Delete</button>`:""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="stock-save">Save</button></footer>`,"narrow");
}
function saveStock(){ const f=$("#stockform"); const fd=Object.fromEntries(new FormData(f)); fd.qty=+fd.qty||0; ["day","week","month"].forEach(k=>{ fd[k] = fd[k]==="" ? null : +fd[k]; });
  if (!fd.name.trim()){ toast("Name is required"); return; }
  if (f.dataset.id){ const s=stockById(f.dataset.id); if (fd.qty < stockOut(s.id)){ toast(`${stockOut(s.id)} are checked out. Total can't go below that.`); return; } Object.assign(s,fd); applyPendingPhoto(s); } else { const s = {id:uid("s"),...fd}; applyPendingPhoto(s); S.stock.push(s); }
  save(); closeModal(); toast("Stock item saved"); rerender(); }

/* ================= Inventory (consumables & parts) ================= */
const invStatus = i => i.qty<=0 ? "Out of stock" : (+i.reorder && i.qty<=+i.reorder) ? "Low stock" : "In stock";
ui.inv = {tab:"active", q:"", sort:"name", cat:""};
function vInventory(){
  const activeInv = S.inventory.filter(i=>!i.retired), retiredInv = S.inventory.filter(i=>i.retired), showRet = ui.inv.tab==="retired";
  const U = ui.inv, nm = i => String(i.name||"").trim().replace(/^[^a-z0-9]+/i,"");
  let list = (showRet ? retiredInv : activeInv).slice();
  if (U.cat) list = list.filter(i=> U.cat==="__none" ? !i.category : (i.category===U.cat || groupOf(i.category)?.parent===U.cat));
  if (U.q){ const q = U.q.toLowerCase(); list = list.filter(i=>[i.name,i.sku,i.vendor,i.category,keptAt(i)].join(" ").toLowerCase().includes(q)); }
  const byGroup = U.sort==="group";
  list.sort((a,b)=> (byGroup ? ((a.category?0:1)-(b.category?0:1) || az(groupLabel(a.category||""), groupLabel(b.category||""))) : 0) || az(nm(a), nm(b)));
  const low = activeInv.filter(i=>invStatus(i)!=="In stock");
  const value = activeInv.reduce((s,i)=>s+(+i.qty||0)*(+i.unitCost||0),0);
  return band("Inventory","Fuel, filters, oil, parts and supplies. Receive stock in, use it on work orders, and get flagged when it runs low.",`<button class="btn" data-act="export-inv">Export to CSV</button><button class="btn primary" data-act="inv-new">+ Add inventory item</button>`) + tabs([["active","Active",activeInv.length],["retired","Retired",retiredInv.length]],ui.inv.tab,"invtab") + `
  <div class="content">
    ${showRet?`<div class="infobox">Retired items are hidden from check-outs and work orders. Open one with Edit to restore it.</div>`:""}
    <div class="kpis" ${showRet?"hidden":""}><div class="kpi"><span class="lbl">Items tracked</span><span class="val num">${activeInv.length}</span></div>
    <div class="kpi ${low.length?"bad":""}"><span class="lbl">Low or out</span><span class="val num">${low.length}</span><span class="note">at or under the reorder point</span></div>
    <div class="kpi"><span class="lbl">Value on hand</span><span class="val num">${money0(value)}</span><span class="note">quantity × unit cost</span></div></div>
    <div class="toolbar" style="margin-bottom:10px;gap:8px;flex-wrap:wrap">
      <div class="search" style="max-width:320px">${ico("search")}<input id="inv-q" type="search" placeholder="Search item, SKU, vendor, location" value="${esc(U.q||"")}" data-input="invq" aria-label="Search inventory"></div>
      <label class="small muted" for="inv-sort">Sort by</label><select id="inv-sort" data-change="invsort"><option value="name" ${byGroup?"":"selected"}>Name (A–Z)</option><option value="group" ${byGroup?"selected":""}>Group, then name</option></select>
      <select data-change="invcat" aria-label="Filter by group"><option value="">All groups</option>${orderedGroups().map(g=>`<option value="${esc(g.name)}" ${U.cat===g.name?"selected":""}>${g.parent?"\u00a0\u00a0\u00a0› ":""}${esc(g.name)}</option>`).join("")}<option value="__none" ${U.cat==="__none"?"selected":""}>No group</option></select>
      ${U.q||U.cat?`<span class="small muted">${list.length} of ${(showRet?retiredInv:activeInv).length} items</span><button class="linkbtn small" data-act="inv-clear">Clear</button>`:""}
    </div>
    <div class="tablewrap"><table><thead><tr><th>Item</th><th>Group</th><th style="text-align:right">On hand</th><th style="text-align:right">Reorder at</th><th>Status</th><th>Checked out</th><th style="text-align:right">Unit cost</th><th style="text-align:right">Charge price</th><th>Vendor</th><th>Last change</th><th></th></tr></thead><tbody>
    ${list.length?list.map((i,ix)=>{ const last=(i.log||[]).slice(-1)[0]; const gh = byGroup && (ix===0 || (list[ix-1].category||"")!==(i.category||"")) ? `<tr class="inv-grp"><td colspan="11">${i.category?esc(groupLabel(i.category)):"No group"} <span class="muted">${list.filter(x=>(x.category||"")===(i.category||"")).length}</span></td></tr>` : ""; return gh + `<tr><td><div class="idcell">${photoThumb(i.photo,"inv",i.id,"inventory")}<div><b>${esc(i.name)}</b> ${canSell(i)?'<span class="chip available">For sale</span>':'<span class="chip neutral">Check-out only</span>'}<div class="small muted">${i.retired?`Retired ${i.retiredAt?fmtDate(i.retiredAt):""}${i.retiredReason?" · "+esc(i.retiredReason):""}`:esc(keptAt(i))}${i.sku?` · <span class="num">${esc(i.sku)}</span>`:""}</div></div></div></td><td class="muted" style="white-space:nowrap" title="${esc(groupLabel(i.category||""))}">${i.category?esc(i.category):"—"}</td><td class="num" style="text-align:right"><b>${i.qty}</b> ${esc(i.unit)}</td><td class="num" style="text-align:right">${+i.reorder?i.reorder:"—"}</td><td>${chipFor(invStatus(i))}</td><td class="small">${S.custody.filter(c=>c.kind==="inv"&&c.invId===i.id&&!c.inAt).map(c=>`${esc(c.holder)} <span class="num muted">×${c.qty}</span>${custodyOverdue(c)?' <span style="color:var(--bad)">late</span>':""}`).join("<br>")||'<span class="muted">—</span>'}</td><td class="num" style="text-align:right">${i.unitCost!=null&&i.unitCost!==""?money(i.unitCost):"—"}</td><td class="num" style="text-align:right">${i.price!=null&&i.price!==""?money(i.price):"—"}</td><td>${esc(i.vendor)||"—"}</td><td class="small">${last?`<span class="num">${last.delta?(last.delta>0?"+":"")+last.delta:""}</span> <span class="muted">${esc(last.reason)} · ${fmtDate(last.at)}</span>`:'<span class="muted">—</span>'}</td><td><div style="display:flex;gap:6px;justify-content:flex-end">${i.retired?"":`<button class="btn sm" data-act="inv-co" data-id="${i.id}" ${+i.qty>0?"":"disabled"}>Check out</button><button class="btn sm" data-act="inv-adj" data-id="${i.id}">Adjust</button>`}<button class="btn sm" data-act="inv-edit" data-id="${i.id}">Edit</button></div></td></tr>`;}).join(""):`<tr><td colspan="11"><div class="empty"><strong>${showRet?"No retired inventory":"No inventory yet"}</strong>${showRet?"Items you retire show up here.":"Add fuel, filters or parts you keep on the shelf."}</div></td></tr>`}
    </tbody></table></div>
    <div class="small muted">Diesel and gasoline were added from your rate sheet's charge prices. Set the gallons on hand and a reorder point to start tracking them.</div>
  </div>`;
}
function invForm(id){
  const i = id ? S.inventory.find(x=>x.id===id) : {name:"",sku:"",unit:"ea",qty:0,reorder:"",unitCost:"",price:"",locationId:invHomeId(),vendor:""};
  openModal(`<header><h2>${id?"Edit inventory item":"Add inventory item"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="invform" class="form" data-id="${id||""}">
    <div class="field full"><label for="iv-name">Name</label><input id="iv-name" name="name" value="${esc(i.name)}" placeholder="Hydraulic filter, Cat 255"></div>
    <div class="field full"><label for="iv-cat">Group</label><select id="iv-cat" name="category">${groupOptions(i.category||"","No group")}</select></div>
    <div class="field"><label for="iv-sku">SKU / part #</label><input id="iv-sku" name="sku" value="${esc(i.sku)}"></div>
    <div class="field"><label for="iv-unit">Unit</label><input id="iv-unit" name="unit" value="${esc(i.unit)}"></div>
    ${id?"":`<div class="field"><label for="iv-qty">Starting quantity</label><input id="iv-qty" type="number" step="0.1" name="qty" value="${i.qty}"></div>`}
    <div class="field"><label for="iv-re">Reorder at</label><input id="iv-re" type="number" step="0.1" name="reorder" value="${i.reorder??""}"></div>
    <div class="field"><label for="iv-cost">Unit cost</label><input id="iv-cost" type="number" step="0.01" name="unitCost" value="${i.unitCost??""}"></div>
    <div class="field"><label for="iv-price">Charge price</label><input id="iv-price" type="number" step="0.01" name="price" value="${i.price??""}"></div>
    <div class="field full"><label class="tgl" for="iv-sell"><input id="iv-sell" type="checkbox" name="sellable" ${canSell(i)?"checked":""}><span class="tgl-ui" aria-hidden="true"></span><span><b>Can be sold</b><span class="hint" style="display:block">On: shows under <b>Inventory sold</b> on rental orders and is billed at the charge price. Off: check-out only. Crews take it and bring it back, and it's never billed.</span></span></label></div>
    <div class="field"><label for="iv-loc">Kept at</label>${keptAtSelect("iv-loc", i.locationId)}</div>
    <div class="field"><label for="iv-vendor">Vendor</label><input id="iv-vendor" name="vendor" value="${esc(i.vendor)}"></div>
    ${photoField(i.photo)}
  </form></div>
  <footer>${id?`<span class="foot-left"><button class="btn danger" data-act="inv-del" data-id="${id}">Delete</button>${i.retired?`<button class="btn" data-act="inv-restore" data-id="${id}">Restore</button>`:`<button class="btn" data-act="inv-retire" data-id="${id}">Retire</button>`}</span>`:""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="inv-save">Save</button></footer>`,"narrow");
}
function saveInv(){ const f=$("#invform"); const fd=Object.fromEntries(new FormData(f));
  if (!fd.name.trim()){ toast("Name is required"); return; }
  const n = v => v===""||v==null?null:+v; fd.reorder=n(fd.reorder); fd.unitCost=n(fd.unitCost); fd.price=n(fd.price); fd.sellable = !!$("#iv-sell")?.checked;
  if (fd.sellable && !(fd.price>0)){ toast("Add a charge price so it can be billed when sold"); $("#iv-price")?.focus(); return; }
  if (f.dataset.id){ const it = S.inventory.find(x=>x.id===f.dataset.id); Object.assign(it, fd); applyPendingPhoto(it); }
  else { fd.qty=+fd.qty||0; const it = {id:uid("i"),...fd,log:fd.qty?[{by:myName(),at:toLocalInput(now()),delta:fd.qty,reason:"Starting count"}]:[]}; applyPendingPhoto(it); S.inventory.push(it); }
  save(); closeModal(); toast("Inventory saved"); rerender(); }
function invAdjust(id){
  const i = S.inventory.find(x=>x.id===id);
  openModal(`<header><h2>Adjust ${esc(i.name)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><div class="photo-row">${i.photo?`<div class="fp-img"><img src="${i.photo}" alt=""></div>`:""}<div class="infobox" style="flex:1">On hand now: <b class="num">${i.qty} ${esc(i.unit)}</b></div></div>
  <form id="adjform" class="form" data-id="${id}">
    <div class="field"><label for="ad-why">Reason</label><select id="ad-why" name="reason"><option>Received</option><option>Used</option><option>Charged to a job</option><option>Count correction</option></select></div>
    <div class="field"><label for="ad-qty">Quantity</label><input id="ad-qty" type="number" step="0.1" name="qty" value="" placeholder="e.g. 50"><span class="hint">Received adds. Used and charged subtract. A correction sets the exact count.</span></div>
    <div class="field full"><label for="ad-note">Note</label><input id="ad-note" name="note" placeholder="PO #, job, who"></div>
  </form>
  <div class="subhead"><span>Recent changes</span></div>
  <div class="lines">${(i.log||[]).slice(-6).reverse().map(l=>`<div class="line"><div>${esc(l.reason)}<div class="calc">${fmtDate(l.at)}</div></div><div class="num">${l.delta?(l.delta>0?"+":"")+l.delta:""}</div><span></span></div>`).join("")||'<div class="line muted">No changes yet.</div>'}</div></div>
  <footer><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="inv-adj-save">Save</button></footer>`,"narrow");
}
function saveAdjust(){ const f=$("#adjform"); const fd=Object.fromEntries(new FormData(f)); const i=S.inventory.find(x=>x.id===f.dataset.id); const q=+fd.qty;
  if (fd.qty===""||isNaN(q)){ toast("Enter a quantity"); return; }
  const delta = fd.reason==="Received" ? q : fd.reason==="Count correction" ? q - i.qty : -q;
  i.qty = +(i.qty + delta).toFixed(2); (i.log ||= []).push({by:myName(),at:toLocalInput(now()),delta:+delta.toFixed(2),reason:fd.reason+(fd.note?` (${fd.note})`:"")});
  save(); closeModal(); toast(`${i.name}: ${i.qty} ${i.unit} on hand`); rerender(); }

/* ================= Groups ================= */
function groupStats(names){
  const its = S.items.filter(i=>names.includes(i.category) && i.status!=="Retired"), st = its.map(itemStatus);
  const billed = S.orders.filter(o=>o.status!=="Cancelled").reduce((s,o)=>s+orderTotals(o).lines.filter(l=>names.includes(itemById(l.itemId)?.category)).reduce((a,l)=>a+l.total,0),0);
  return {n:its.length,avail:st.filter(s=>s==="Available").length,out:st.filter(s=>s==="Rented Out"||s==="Overdue").length,co:st.filter(s=>s==="Checked Out").length,stock:S.stock.filter(s=>names.includes(s.category)).length,inv:S.inventory.filter(i=>!i.retired&&names.includes(i.category)).length,u:its.length?its.reduce((s,i)=>s+utilization(i.id),0)/its.length:0,billed,list:its.reduce((s,i)=>s+(+i.month||0),0)};
}
function vGroups(){
  const open = ui.groups.open && groupOf(ui.groups.open) && subGroups(ui.groups.open).length ? ui.groups.open : null;
  if (!open) ui.groups.open = null;
  const list = open ? subGroups(open) : orderedGroups().filter(g=>!(g.parent && groupOf(g.parent)));
  const rows = list.map(g=>{ const kids = open ? 0 : subGroups(g.name).length;
    return {g, kids, ...groupStats(kids ? groupFamily(g.name) : [g.name])}; });
  const own = open ? S.items.filter(i=>i.category===open && i.status!=="Retired").length : 0;
  if (own) rows.push({g:{name:open,desc:"Assets in this group that aren't in a sub-group"}, own:true, label:`Other ${open} assets`, ...groupStats([open])});
  const pg = open ? groupOf(open) : null;
  const head = open
    ? band(open, `<button class="linkbtn" data-act="group-close">‹ All groups</button> · ${subGroups(open).length} sub-group${subGroups(open).length>1?"s":""}${pg.desc?" · "+esc(pg.desc):""}. Click a sub-group to see its assets.`, `<button class="btn" data-act="group-edit" data-v="${esc(open)}">Edit ${esc(open)}</button><button class="btn primary" data-act="group-new" data-v="${esc(open)}">+ Add sub-group</button>`)
    : band("Groups","How the fleet is organized. Click a group with sub-groups to open it, or any other group to see its assets.",`<button class="btn" data-act="group-combine">Combine groups</button><button class="btn primary" data-act="group-new">+ Add group</button>`);
  return head + `
  <div class="content"><div class="tablewrap"><table><thead><tr><th>${open?"Sub-group":"Group"}</th><th style="text-align:right">Assets</th><th style="text-align:right">Available</th><th style="text-align:right">On rent</th><th style="text-align:right">Checked out</th><th style="text-align:right">Stock items</th><th style="text-align:right">Inventory items</th><th>Use (90 days)</th><th style="text-align:right">Month list value</th><th style="text-align:right">Billed to date</th><th></th></tr></thead><tbody>
  ${rows.map(r=>{ const act = r.kids ? `data-act="group-open" data-v="${esc(r.g.name)}"` : r.n ? `data-act="group-assets" data-v="${esc(r.g.name)}"${r.own?' data-exact="1"':""}` : "";
    return `<tr class="grp-row ${act?"clickable":""} ${r.own?"own-row":""}" ${act}><td><b class="grp-name">${r.kids?ico("groups"):""}${esc(r.label||r.g.name)}</b>${r.kids?`<div class="small grp-kids">${r.kids} sub-group${r.kids>1?"s":""} ›</div>`:""}${r.g.desc?`<div class="small muted">${esc(r.g.desc)}</div>`:""}</td><td class="num" style="text-align:right">${r.n}</td><td class="num" style="text-align:right">${r.avail}</td><td class="num" style="text-align:right">${r.out}</td><td class="num" style="text-align:right">${r.co}</td><td class="num" style="text-align:right">${r.stock||"—"}</td><td class="num" style="text-align:right">${r.inv||"—"}</td>
  <td><div class="bar-row" style="grid-template-columns:90px 40px"><div class="track"><div class="fill" style="width:${r.u*100}%"></div></div><span class="v small">${Math.round(r.u*100)}%</span></div></td>
  <td class="num" style="text-align:right">${money0(r.list)}</td><td class="num" style="text-align:right">${money(r.billed)}</td>
  <td><div style="display:flex;gap:6px;justify-content:flex-end">${r.kids?`<button class="btn sm" data-act="group-open" data-v="${esc(r.g.name)}">Open</button>`:`<button class="btn sm" data-act="group-assets" data-v="${esc(r.g.name)}"${r.own?' data-exact="1"':""} ${r.n?"":"disabled"}>Assets</button>`}${r.own?"":`<button class="btn sm" data-act="group-edit" data-v="${esc(r.g.name)}">Edit</button>`}</div></td></tr>`; }).join("") || `<tr><td colspan="11"><div class="empty">No groups yet.</div></td></tr>`}
  </tbody></table></div></div>`;
}
function groupForm(name, parent){
  const g = name ? S.groups.find(x=>x.name===name) : {name:"",desc:"",parent:parent||""};
  const inUse = name && (S.items.some(i=>i.category===name) || S.stock.some(s=>s.category===name) || S.inventory.some(i=>i.category===name) || subGroups(name).length);
  const hasKids = name && subGroups(name).length;
  const parents = S.groups.filter(x=>x.name!==name && !(x.parent && groupOf(x.parent))).sort((a,b)=>byNum(a.name,b.name));
  openModal(`<header><h2>${name?"Edit group":"Add group"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="groupform" class="form" data-orig="${esc(name||"")}">
    <div class="field full"><label for="gp-name">Name</label><input id="gp-name" name="name" value="${esc(g.name)}"><span class="hint">${name?"Renaming moves every asset in this group with it.":""}</span></div>
    <div class="field full"><label for="gp-parent">Sub-group of</label><select id="gp-parent" name="parent" ${hasKids?"disabled":""}><option value="">None (top-level group)</option>${parents.map(p=>`<option ${g.parent===p.name?"selected":""}>${esc(p.name)}</option>`).join("")}</select><span class="hint">${hasKids?`This group has ${subGroups(name).length} sub-group${subGroups(name).length>1?"s":""}, so it stays top-level.`:"Pick a parent to make this a sub-group, for example Skid Steers under 100 - Excavation."}</span></div>
    <div class="field full"><label for="gp-desc">Description</label><input id="gp-desc" name="desc" value="${esc(g.desc)}"></div>
  </form></div>
  <footer>${name?`<span class="foot-left"><button class="btn danger" data-act="group-del" data-v="${esc(name)}" ${inUse?'disabled title="Move its assets, stock, inventory and sub-groups first, or combine it into another group"':""}>Delete</button><button class="btn" data-act="group-combine" data-v="${esc(name)}">Combine…</button>${g.parent?`<button class="btn" data-act="group-ungroup" data-v="${esc(name)}">Make top-level</button>`:""}</span>`:""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="group-save">Save</button></footer>`,"narrow");
}
let cmb = null;
function combineForm(pre){
  cmb = {picked: pre ? [pre] : [], mode:"parent", parent:"__new", newName:"", into:""};
  drawCombine();
}
function drawCombine(){
  const all = orderedGroups(), c = cmb;
  const tops = S.groups.filter(g=>!(g.parent && groupOf(g.parent))).sort((a,b)=>byNum(a.name,b.name));
  openModal(`<header><h2>Combine groups</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody">
  <div class="field"><label>Groups to combine (${c.picked.length} picked)</label><div class="picker" style="max-height:260px">${all.map(g=>{ const n = S.items.filter(i=>i.category===g.name).length; return `<label class="pick" style="grid-template-columns:24px minmax(0,1fr) auto"><input type="checkbox" data-act="cmb-pick" value="${esc(g.name)}" ${c.picked.includes(g.name)?"checked":""}><span>${g.parent&&groupOf(g.parent)?`<span class="muted">${esc(g.parent)} › </span>`:""}${esc(g.name)}</span><span class="small muted num">${n} asset${n===1?"":"s"}</span></label>`; }).join("")}</div></div>
  <div class="field"><label>How to combine</label>
    <label class="check"><input type="radio" name="cmb-mode" data-act="cmb-mode" value="parent" ${c.mode==="parent"?"checked":""}> Put them under one parent group. Each keeps its own assets and is listed under the parent.</label>
    <label class="check"><input type="radio" name="cmb-mode" data-act="cmb-mode" value="merge" ${c.mode==="merge"?"checked":""}> Merge them into a single group. Their assets move and the other groups are removed.</label>
  </div>
  ${c.mode==="parent" ? `
  <div class="form">
    <div class="field"><label for="cmb-parent">Parent group</label><select id="cmb-parent" data-cmb="parent"><option value="__new" ${c.parent==="__new"?"selected":""}>New parent group…</option>${tops.map(g=>`<option value="${esc(g.name)}" ${c.parent===g.name?"selected":""}>${esc(g.name)}</option>`).join("")}</select></div>
    ${c.parent==="__new"?`<div class="field"><label for="cmb-new">New parent name</label><input id="cmb-new" data-cmb="newName" value="${esc(c.newName)}" placeholder="Earthmoving Equipment"></div>`:""}
  </div>
  <div class="small muted">${c.picked.length?`The parent will list: ${c.picked.filter(n=>n!==c.parent).map(n=>`<b>${esc(n)}</b>`).join(", ") || "(pick at least one group besides the parent)"}`:"Tick the groups that belong together."}</div>` : `
  <div class="field"><label for="cmb-into">Merge into</label><select id="cmb-into" data-cmb="into">${groupOptions(c.into, "Choose the group to keep")}</select></div>`}
  </div><footer><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="group-combine-save">${c.mode==="parent"?"Combine under parent":"Merge groups"}</button></footer>`, "narrow");
}
function saveCombine(){
  const c = cmb;
  if (c.mode==="parent"){
    const parentName = c.parent==="__new" ? (c.newName||"").trim() : c.parent;
    if (!parentName){ toast("Name the new parent group"); return; }
    if (c.parent==="__new" && groupOf(parentName)){ toast(`A group called ${parentName} already exists. Pick it from the list instead.`); return; }
    const kids = c.picked.filter(n=>n!==parentName);
    if (!kids.length){ toast("Tick at least one group to put under the parent"); return; }
    if (!groupOf(parentName)) S.groups.push({name:parentName, desc:""});
    const p = groupOf(parentName); delete p.parent;
    kids.forEach(n=>{ subGroups(n).forEach(s=>s.parent = parentName); const g = groupOf(n); if (g) g.parent = parentName; });
    cmb = null; save(); closeModal(); rerender(); toast(`${parentName} now lists ${kids.length} group${kids.length>1?"s":""}: ${kids.join(", ")}`);
    return;
  }
  const into = c.into, picked = c.picked.filter(n=>n!==into);
  if (!into){ toast("Choose the group to keep"); return; }
  if (!picked.length){ toast("Tick at least one other group to merge into it"); return; }
  const target = groupOf(into); let assets = 0, stock = 0;
  picked.forEach(n=>{
    S.items.forEach(i=>{ if (i.category===n){ i.category = into; itemLog(i, `Group ${n} → ${into} (merged)`); assets++; } });
    S.stock.forEach(s=>{ if (s.category===n){ s.category = into; stock++; } });
    S.inventory.forEach(i=>{ if (i.category===n){ i.category = into; stock++; } });
    S.groups.forEach(g=>{ if (g.parent===n) g.parent = target.parent ? target.parent : into; });
    if (ui.items.cat===n) ui.items.cat = into;
  });
  S.groups = S.groups.filter(g=>!picked.includes(g.name));
  S.groups.forEach(g=>{ if (g.parent===g.name || (g.parent && !groupOf(g.parent))) delete g.parent; });
  cmb = null; save(); closeModal(); rerender(); toast(`Merged ${picked.length} group${picked.length>1?"s":""} into ${into} (${assets} asset${assets===1?"":"s"}${stock?`, ${stock} stock and inventory item${stock===1?"":"s"}`:""} moved)`);
}
function saveGroup(){ const f=$("#groupform"); const fd=Object.fromEntries(new FormData(f)); const orig=f.dataset.orig; const nm=fd.name.trim();
  if (!nm){ toast("Name is required"); return; }
  if (nm!==orig && S.groups.some(g=>g.name===nm)){ toast(`A group called ${nm} already exists`); return; }
  const parent = fd.parent || "";
  if (orig){ const g=S.groups.find(x=>x.name===orig); g.name=nm; g.desc=fd.desc; if (!subGroups(orig).length) g.parent = parent; else delete g.parent;
    S.groups.forEach(x=>{ if (x.parent===orig) x.parent = nm; });
    S.items.forEach(i=>{ if(i.category===orig) i.category=nm; }); S.stock.forEach(s=>{ if(s.category===orig) s.category=nm; }); S.inventory.forEach(i=>{ if(i.category===orig) i.category=nm; }); if (ui.items.cat===orig) ui.items.cat=nm; }
  else S.groups.push({name:nm,desc:fd.desc,...(parent?{parent}:{})});
  if (!S.groups.find(x=>x.name===nm).parent) delete S.groups.find(x=>x.name===nm).parent;
  save(); closeModal(); toast("Group saved"); rerender(); }

/* ================= Users, roles and sign-in ================= */
const ROLES = {
  "Admin":            {perms:["admin","notes","orders","checkout","assets","shop","people","reports","setup","users"], desc:"Everything, including users, rates and resetting data."},
  "Manager":          {perms:["notes","orders","checkout","assets","shop","people","reports","setup"], desc:"Runs the operation: rentals, fleet, shop, rates and reports. Can't manage users."},
  "Field Crew":       {perms:["notes","checkout"], desc:"Checks tools and equipment out and back in. Sees everything else read-only."},
  "Viewer":           {perms:[], desc:"Read-only access to orders, assets and the shop board. Can't comment."},
  "Superintendent":   {perms:["request"], desc:"Sees only equipment that's ready to rent and all inventory. Requests what their job needs; admins approve."},
};
const PERM_LABELS = {request:"Request equipment & inventory",admin:"Archive & delete customers",notes:"Comments & files",orders:"Rentals & orders",checkout:"Check in / out",assets:"Edit assets, stock, groups",shop:"Maintenance & inventory",people:"Customers & locations",reports:"Reports",setup:"Rates & data file",users:"Users & roles"};
const ACT_PERM = {};
const mapPerm = (p, acts) => acts.split(" ").forEach(a=>ACT_PERM[a]=p);
mapPerm("orders","new-order rent-item edit-order dup-order save-order mark-out req-pickup req-pickup2 cancel-order checkin ci-save ci-quick bulk add-charge add-sale del-sale cycle-mark cycle-void");
mapPerm("checkout","co-undo co-new co-item co-save co-in co-bulkin coin-save stock-co co-add co-line-del inv-co");
mapPerm("assets","group-combine group-combine-save a-retire a-restore a-hours-save a-photo-add a-photo-del a-photo-primary dup-item a-hours new-item edit-item save-item del-item stock-new stock-edit stock-save stock-del group-new group-edit group-save group-del");
mapPerm("request","rq-add rq-remove rq-inc rq-dec rq-clear rq-review rq-line-del rq-send rq-cancel");
mapPerm("admin","rq-decline rq-approve rq-decline-save rq-approve-save rq-order");
mapPerm("shop","sr-accept sr-stage sr-qc-open sr-qc-pass sr-qc-fail sr-reopen sr-task-new sr-task-add sr-task-del sr-labor-new sr-labor-add sr-labor-del sr-part-new sr-part-add sr-part-del sr-addfile inv-retire inv-restore m-addfile m-file-del new-maint edit-maint save-maint complete-maint del-maint inv-new inv-edit inv-save inv-adj inv-adj-save inv-del m-addpart");
mapPerm("people","new-cust edit-cust new-loc edit-loc save-simple");
mapPerm("setup","logo-up logo-del add-rate del-rate file-open reset");
mapPerm("notes","a-comment a-file-add a-file-del");
mapPerm("reports","rpt-csv rpt-xlsx rpt-all");
mapPerm("admin","cust-archive cust-restore cust-del");
mapPerm("users","user-new user-edit user-save user-del");
const PAGE_PERM = {setup:"setup", users:"users", requests:"admin", request:"request", "my-requests":"request", "my-reports":"request"};

let SESSION_ID = null;
try { SESSION_ID = localStorage.getItem(KEY+"-session"); } catch(e) {}
function me(){ const u = S.users.find(x=>x.id===SESSION_ID && x.active); return u || null; }
function myName(){ return me()?.name || ""; }
function can(p){ const u = me(); return !!u && ROLES[u.role]?.perms.includes(p); }
function setSession(id){ SESSION_ID = id; try { id ? localStorage.setItem(KEY+"-session", id) : localStorage.removeItem(KEY+"-session"); } catch(e) {} }
async function hashPw(id, pw){
  const str = "bh-rentals:" + id + ":" + pw;
  try { const h = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str)); return [...new Uint8Array(h)].map(b=>b.toString(16).padStart(2,"0")).join(""); }
  catch(e) { let h1 = 0x811c9dc5, h2 = 0; for (let i=0;i<str.length;i++){ h1 ^= str.charCodeAt(i); h1 = Math.imul(h1,16777619)>>>0; h2 = (h2*31 + str.charCodeAt(i))>>>0; } return "f" + h1.toString(16) + h2.toString(16); }
}
function applyPerms(root){
  root.querySelectorAll("[data-act]").forEach(el=>{ const p = ACT_PERM[el.dataset.act]; if (p) el.hidden = !can(p); });
  if (!can("setup")) root.querySelectorAll("[data-set],[data-rate]").forEach(el=>el.disabled = true);
}

/* ---- Sign-in screens ---- */
function authScreen(mode, msg=""){
  const first = S.users.find(u=>u.role==="Admin");
  const card = body => `<div class="auth"><div class="auth-card"><div class="logo auth-logo ${S.settings.logo?"has-img":""}">${brandHtml("auth")}</div>${body}</div><p class="auth-foot">${esc(S.settings.company)} rentals · ${esc(S.settings.phone)}</p></div>`;
  if (mode==="setup") return card(`<h1>Set up your admin account</h1><p class="muted">No one has signed in yet. Create the administrator's password to get started. You can add the rest of your team after that.</p>
    <form id="auth-form" data-mode="setup" class="form" style="grid-template-columns:1fr">
      <div class="field"><label for="au-name">Name</label><input id="au-name" name="name" value="${esc(first?.name||"")}" required></div>
      <div class="field"><label for="au-email">Email</label><input id="au-email" name="email" type="email" value="${esc(first?.email||"")}" required autocomplete="username"></div>
      <div class="field"><label for="au-pw">Password</label><input id="au-pw" name="pw" type="password" minlength="6" required autocomplete="new-password"><span class="hint">At least 6 characters</span></div>
      <div class="field"><label for="au-pw2">Confirm password</label><input id="au-pw2" name="pw2" type="password" required autocomplete="new-password"></div>
      ${msg?`<div class="auth-err">${esc(msg)}</div>`:""}
      <button class="btn primary auth-btn" type="submit">Create admin account</button></form>
    <div class="auth-file"><span class="small muted">Already have a data file?</span><button class="btn" type="button" data-act="file-open">Open data file…</button></div>`);
  if (mode==="change") return card(`<h1>Choose a new password</h1><p class="muted">Hi ${esc(me()?.name)}. Your password was set by an administrator. Pick your own before you continue.</p>
    <form id="auth-form" data-mode="change" class="form" style="grid-template-columns:1fr">
      <div class="field"><label for="au-pw">New password</label><input id="au-pw" name="pw" type="password" minlength="6" required autocomplete="new-password"></div>
      <div class="field"><label for="au-pw2">Confirm password</label><input id="au-pw2" name="pw2" type="password" required autocomplete="new-password"></div>
      ${msg?`<div class="auth-err">${esc(msg)}</div>`:""}
      <button class="btn primary auth-btn" type="submit">Save and continue</button>
      <button class="btn" type="button" data-act="logout">Sign out</button></form>`);
  return card(`<h1>Sign in</h1><p class="muted">Use the email and password your administrator gave you.</p>
    <form id="auth-form" data-mode="login" class="form" style="grid-template-columns:1fr">
      <div class="field"><label for="au-email">Email</label><input id="au-email" name="email" type="email" required autocomplete="username"></div>
      <div class="field"><label for="au-pw">Password</label><input id="au-pw" name="pw" type="password" required autocomplete="current-password"></div>
      ${msg?`<div class="auth-err">${esc(msg)}</div>`:""}
      <button class="btn primary auth-btn" type="submit">Sign in</button></form>
    <p class="small muted" style="margin:0">Forgot your password? Ask an administrator to reset it from <b>Users</b>.</p>
    <div class="auth-file"><span class="small muted">Using your team's data file from OneDrive?</span><button class="btn" type="button" data-act="file-open">Open data file…</button></div>`);
}
function showAuth(mode, msg){ $("#auth-root").innerHTML = authScreen(mode, msg); document.querySelector("#auth-root input")?.focus(); }
function authGate(){
  const u = me();
  if (!S.users.some(x=>x.hash)) { $(".app").hidden = true; showAuth("setup"); return false; }
  if (!u) { $(".app").hidden = true; showAuth("login"); return false; }
  if (u.mustChange) { $(".app").hidden = true; showAuth("change"); return false; }
  $("#auth-root").innerHTML = ""; $(".app").hidden = false; return true;
}
document.addEventListener("submit", async e => {
  if (e.target.id !== "auth-form") return;
  e.preventDefault();
  const f = e.target, fd = Object.fromEntries(new FormData(f)), mode = f.dataset.mode;
  if (mode==="login"){
    const u = S.users.find(x=>x.email.toLowerCase()===fd.email.trim().toLowerCase());
    if (!u || !u.hash || (await hashPw(u.id, fd.pw)) !== u.hash) { showAuth("login","That email and password don't match. Check both and try again."); $("#au-email").value = fd.email; return; }
    if (!u.active) { showAuth("login","This account is turned off. Ask an administrator to turn it back on."); return; }
    setSession(u.id); u.lastLogin = toLocalInput(now()); save(); render(); return;
  }
  if (fd.pw.length < 6) { showAuth(mode,"Use at least 6 characters."); return; }
  if (fd.pw !== fd.pw2) { showAuth(mode,"The two passwords don't match."); return; }
  if (mode==="setup"){
    const u = S.users.find(x=>x.role==="Admin") || (S.users.push({id:uid("u-"),role:"Admin",active:true}), S.users[S.users.length-1]);
    u.name = fd.name.trim(); u.email = fd.email.trim(); u.hash = await hashPw(u.id, fd.pw); u.mustChange = false; u.active = true;
    setSession(u.id); u.lastLogin = toLocalInput(now()); save(); render(); toast("Admin account created"); return;
  }
  if (mode==="change"){ const u = me(); u.hash = await hashPw(u.id, fd.pw); u.mustChange = false; save(); render(); toast("Password updated"); }
});

/* ---- Profile menu (the initials badge) ---- */
function toggleProfileMenu(){
  let m = $("#pmenu"); if (m) { m.remove(); return; }
  const u = me(); if (!u) return;
  m = document.createElement("div"); m.id = "pmenu"; m.className = "pmenu"; m.setAttribute("role","menu");
  m.innerHTML = `<div class="pm-head">${avatarHtml(u,"big")}<div style="min-width:0"><b>${esc(u.name)}</b><div class="small muted trunc" style="max-width:190px">${esc(u.email)}</div>${chipFor(u.role)}</div></div>
    <button role="menuitem" data-act="profile">Edit profile</button>
    <button role="menuitem" data-act="profile" data-pw="1">Change password</button>
    ${can("users")?`<button role="menuitem" data-act="goto" data-to="users">Manage users</button>`:""}
    <button role="menuitem" data-act="theme">Switch light / dark</button>
    <button role="menuitem" data-act="logout" class="pm-out">Sign out</button>`;
  document.body.appendChild(m);
  const r = $("#avatar-btn").getBoundingClientRect(); const narrow = matchMedia("(max-width:760px)").matches;
  m.style.left = (narrow ? Math.max(8, r.left) : r.right + 8) + "px"; m.style.top = (narrow ? r.bottom + 6 : r.top) + "px";
}
function profileForm(focusPw){
  const u = me();
  openModal(`<header><h2>Your profile</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="profileform" class="form">
    <div class="field full photo-row">${avatarHtml(u,"xl")}<div style="display:flex;flex-direction:column;gap:6px"><b>${esc(u.name)}</b><div class="small muted">${esc(u.role)} · ${(ROLES[u.role]?.desc||"")}</div><div style="display:flex;gap:6px"><button type="button" class="btn sm" data-act="u-photo" data-id="${u.id}">${u.photo?"Change photo":"Upload photo"}</button>${u.photo?`<button type="button" class="btn sm" data-act="u-photo-del" data-id="${u.id}">Remove</button>`:""}</div></div></div>
    <div class="field"><label for="pf-name">Name</label><input id="pf-name" name="name" value="${esc(u.name)}"></div>
    <div class="field"><label for="pf-email">Email (sign-in)</label><input id="pf-email" name="email" type="email" value="${esc(u.email)}"></div>
    <div class="field"><label for="pf-phone">Phone</label><input id="pf-phone" name="phone" value="${esc(u.phone||"")}"></div>
    <div class="field"><label for="pf-title">Job title</label><input id="pf-title" name="title" value="${esc(u.title||"")}"></div>
    <div class="field full subhead" style="margin:6px 0 0"><span>Change password</span><span class="small" style="text-transform:none;letter-spacing:0">Leave blank to keep your current one</span></div>
    <div class="field"><label for="pf-cur">Current password</label><input id="pf-cur" name="cur" type="password" autocomplete="current-password"></div>
    <div class="field"><label for="pf-pw">New password</label><input id="pf-pw" name="pw" type="password" autocomplete="new-password"></div>
    <div class="field"><label for="pf-pw2">Confirm new password</label><input id="pf-pw2" name="pw2" type="password" autocomplete="new-password"></div>
  </form><div id="pf-err"></div></div>
  <footer><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="profile-save">Save profile</button></footer>`,"narrow");
  if (focusPw) $("#pf-cur").focus();
}
async function saveProfile(){
  const u = me(), fd = Object.fromEntries(new FormData($("#profileform"))), err = m => $("#pf-err").innerHTML = `<div class="auth-err">${esc(m)}</div>`;
  if (!fd.name.trim() || !fd.email.trim()) return err("Name and email are required.");
  if (S.users.some(x=>x.id!==u.id && x.email.toLowerCase()===fd.email.trim().toLowerCase())) return err("Another user already signs in with that email.");
  if (fd.pw || fd.cur){
    if ((await hashPw(u.id, fd.cur)) !== u.hash) return err("Your current password isn't right.");
    if (fd.pw.length < 6) return err("Use at least 6 characters for the new password.");
    if (fd.pw !== fd.pw2) return err("The new passwords don't match.");
    u.hash = await hashPw(u.id, fd.pw);
  }
  Object.assign(u, {name:fd.name.trim(), email:fd.email.trim(), phone:fd.phone, title:fd.title});
  save(); closeModal(); renderRail(); toast(fd.pw ? "Profile and password saved" : "Profile saved");
}

/* ---- Users page (Admin) ---- */
function vUsers(){
  const rows = S.users.slice().sort((a,b)=>a.name.localeCompare(b.name));
  const perms = Object.keys(PERM_LABELS);
  return band("Users","Who can sign in and what each person can do.",`<button class="btn primary" data-act="user-new">+ Add user</button>`) + `
  <div class="content">
    <div class="tablewrap"><table><thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Last sign-in</th><th></th></tr></thead><tbody>
    ${rows.map(u=>`<tr><td><div style="display:flex;gap:10px;align-items:center">${avatarHtml(u,"sm")}<div><b>${esc(u.name)}</b>${u.id===SESSION_ID?' <span class="small muted">(you)</span>':""}${u.title?`<div class="small muted">${esc(u.title)}</div>`:""}</div></div></td>
      <td>${esc(u.email)}</td><td>${chipFor(u.role)}</td>
      <td>${!u.active?chipFor("Disabled"):!u.hash?chipFor("No password"):u.mustChange?chipFor("Temporary password"):chipFor("Active")}</td>
      <td class="num">${u.lastLogin?fmtDT(u.lastLogin):'<span class="muted">Never</span>'}</td>
      <td><button class="btn sm" data-act="user-edit" data-id="${u.id}">Edit</button></td></tr>`).join("")}
    </tbody></table></div>
    <div class="panel"><h3>What each role can do</h3><div class="tablewrap" style="border:0"><table style="min-width:760px"><thead><tr><th>Role</th>${perms.map(p=>`<th style="white-space:normal;text-transform:none;letter-spacing:0;font-size:11.5px">${PERM_LABELS[p]}</th>`).join("")}</tr></thead><tbody>
    ${Object.entries(ROLES).map(([r,v])=>`<tr><td><b>${r}</b><div class="small muted" style="max-width:240px">${v.desc}</div></td>${perms.map(p=>`<td style="text-align:center">${v.perms.includes(p)?'<span class="yes" aria-label="Yes">✓</span>':'<span class="muted" aria-label="No">—</span>'}</td>`).join("")}</tr>`).join("")}
    </tbody></table></div><div class="body small muted">Everyone who can sign in can view orders, assets, check-outs and the shop board. Buttons for actions a role can't take are hidden.</div></div>
  </div>`;
}
function userForm(id){
  const u = id ? S.users.find(x=>x.id===id) : {name:"",email:"",role:"Field Crew",active:true,title:"",phone:""};
  const self = id===SESSION_ID;
  openModal(`<header><h2>${id?"Edit user":"Add user"}</h2><button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody"><form id="userform" class="form" data-id="${id||""}">
    ${id?`<div class="field full photo-row">${avatarHtml(u,"xl")}<div style="display:flex;gap:6px;flex-wrap:wrap"><button type="button" class="btn sm" data-act="u-photo" data-id="${id}">${u.photo?"Change photo":"Upload photo"}</button>${u.photo?`<button type="button" class="btn sm" data-act="u-photo-del" data-id="${id}">Remove</button>`:""}</div></div>`:""}
    <div class="field"><label for="us-name">Name</label><input id="us-name" name="name" value="${esc(u.name)}"></div>
    <div class="field"><label for="us-email">Email (sign-in)</label><input id="us-email" name="email" type="email" value="${esc(u.email)}"></div>
    <div class="field"><label for="us-title">Job title</label><input id="us-title" name="title" value="${esc(u.title||"")}"></div>
    <div class="field"><label for="us-phone">Phone</label><input id="us-phone" name="phone" value="${esc(u.phone||"")}"></div>
    <div class="field full"><label for="us-role">Role</label><select id="us-role" name="role" ${self?"disabled":""}>${Object.keys(ROLES).map(r=>`<option ${r===u.role?"selected":""}>${r}</option>`).join("")}</select><span class="hint" id="us-role-desc">${self?"You can't change your own role.":(ROLES[u.role]?.desc||"")}</span></div>
    <div class="field full" id="us-cust-wrap" ${u.role==="Superintendent"?"":"hidden"}><label for="us-cust">Customer record (for their reports)</label><select id="us-cust" name="customerId"><option value="">Match by name${superCustomer(u)&&!u.customerId?` (${esc(superCustomer(u).name)})`:""}</option>${byName(S.customers.filter(c=>!c.deleted)).map(c=>`<option value="${c.id}" ${u.customerId===c.id?"selected":""}>${esc(c.name)}${c.position?` · ${esc(c.position)}`:""}</option>`).join("")}</select><span class="hint">Superintendents only see orders for this customer</span></div>
    <div class="field"><label for="us-pw">${id?"Reset password":"Temporary password"}</label><input id="us-pw" name="pw" type="text" autocomplete="off" placeholder="${id?"Leave blank to keep":"At least 6 characters"}"><span class="hint">They'll be asked to pick their own after signing in.</span></div>
    <div class="field"><label for="us-active">Access</label><select id="us-active" name="active" ${self?"disabled":""}><option value="1" ${u.active?"selected":""}>Can sign in</option><option value="0" ${u.active?"":"selected"}>Turned off</option></select></div>
  </form><div id="us-err"></div></div>
  <footer>${id&&!self?`<button class="btn danger" data-act="user-del" data-id="${id}" style="margin-right:auto">Delete user</button>`:""}<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="user-save">${id?"Save user":"Add user"}</button></footer>`,"narrow");
}
async function saveUser(){
  const f = $("#userform"), fd = Object.fromEntries(new FormData(f)), id = f.dataset.id, err = m => $("#us-err").innerHTML = `<div class="auth-err">${esc(m)}</div>`;
  const u = id ? S.users.find(x=>x.id===id) : null;
  if (!fd.name.trim() || !fd.email.trim()) return err("Name and email are required.");
  if (S.users.some(x=>x.id!==id && x.email.toLowerCase()===fd.email.trim().toLowerCase())) return err("Another user already signs in with that email.");
  if (!id && (fd.pw||"").length < 6) return err("Give them a temporary password of at least 6 characters.");
  if (id && fd.pw && fd.pw.length < 6) return err("Use at least 6 characters.");
  const role = id===SESSION_ID ? u.role : fd.role, active = id===SESSION_ID ? true : fd.active!=="0";
  const admins = S.users.filter(x=>x.role==="Admin" && x.active && x.id!==id).length;
  if (!admins && (role!=="Admin" || !active)) return err("Keep at least one active Admin so someone can manage users.");
  const rec = u || {id:uid("u-")};
  Object.assign(rec, {name:fd.name.trim(), email:fd.email.trim(), title:fd.title, phone:fd.phone, role, active});
  if (role==="Superintendent" && fd.customerId) rec.customerId = fd.customerId; else delete rec.customerId;
  if (fd.pw){ rec.hash = await hashPw(rec.id, fd.pw); rec.mustChange = true; }
  if (!u) S.users.push(rec);
  save(); closeModal(); rerender(); toast(u ? "User saved" : `${rec.name} added. Share their email and temporary password with them.`);
}

/* ================= Printable invoice ================= */
const IN_VIEWER = !!(window.claude && typeof window.claude.use === "function");
function termsDays(t){ const m = /Net\s*(\d+)/i.exec(t||""); return m ? +m[1] : 0; }
function invoiceData(o, scope={}){
  const n = toLocalInput(now()), cyc = !!o.cycle, cn = cyc ? (scope.cycle||0) : 0;
  const final = o.status==="Returned" && !cn, bill = cn ? billFor(o, cn) : null;
  let through, from = o.out, pieces = null;
  if (cn){ const sg = bill || cycleSegs(o, cycEnd(o))[cn-1]; pieces = [{from:sg.from, to:sg.to, len:sg.len || rentalDays(sg.from, toLocalInput(cycleAt(o, cn)))}]; from = sg.from; through = sg.to; }
  else if (cyc && final){ through = orderRentalEnd(o); pieces = finalPieces(o, through); from = pieces[0]?.from || through; }
  else through = final ? orderRentalEnd(o) : (isOverdue(o) ? n : o.ret);
  const keys = cyc && (cn || final) ? (bill ? bill.keys||[] : unbilledKeys(o)) : null, has = k => !keys || keys.includes(k);
  const g = +S.settings.taxRate || 0;
  const lines = (pieces && !pieces.length) ? [] : o.lines.map(l=>{ const it = itemById(l.itemId), p = pieces ? piecesRental(l, pieces) : priceLine(l, o.out, through);
    const rates = [l.day!=null&&l.day!==""?money(l.day)+"/day":"", l.week!=null&&l.week!==""?money(l.week)+"/wk":"", l.month!=null&&l.month!==""?money(l.month)+"/mo":""].filter(Boolean).join(" · ");
    const tr = it && it.taxRate!=null && it.taxRate!=="" ? +it.taxRate : g;
    return {id:l.itemId, name:it?.name||"(asset removed)", serial:it?.serial||"", days:p.days, label:p.label, rates, total:p.total, tax:p.total*tr/100}; });
  const rental = lines.reduce((s,l)=>s+l.total,0);
  const bundleDisc = o.bundle ? lines.filter(l=>o.bundle.itemIds.includes(l.id)).reduce((s,l)=>s+l.total,0)*(+o.bundle.pct||0)/100 : 0;
  const delivery = has("delivery") ? (+o.delivery?.rate||0) * (+o.delivery?.trips||0) : 0;
  const charges = (o.charges||[]).map((c,i)=>({i, name:c.name, qty:+c.qty||0, rate:+c.rate||0, total:(+c.qty||0)*(+c.rate||0)})).filter(c=>has("charge:"+c.i));
  const chargesTotal = charges.reduce((s,c)=>s+c.total,0), adj = has("adj") ? +o.adjustment||0 : 0;
  const sales = (o.sales||[]).map((x,i)=>{ const it = invById(x.invId); return {i, name:it?.name||"Inventory item", sku:it?.sku||"", unit:it?.unit||"", qty:+x.qty||0, price:+x.price||0, total:(+x.qty||0)*(+x.price||0)}; }).filter(x=>has("sale:"+x.i));
  const salesSum = sales.reduce((s,x)=>s+x.total,0);
  const subtotal = rental - bundleDisc + delivery + chargesTotal + salesSum + adj;
  const tax = lines.reduce((s,l)=>s+l.tax,0) * (rental ? (rental-bundleDisc)/rental : 1) + (delivery + chargesTotal + salesSum) * g / 100;
  const invDate = bill ? bill.at.slice(0,10) : n.slice(0,10), days = termsDays(S.settings.terms);
  const due = days ? toLocalInput(new Date(parseD(invDate+"T00:00").getTime() + days*DAY)).slice(0,10) : invDate;
  const prior = cyc ? (o.bills||[]).filter(b=>!cn || b.n < cn) : [];
  const invNo = cn ? `INV-${o.no}-C${cn}` : `INV-${o.no}${final||cyc?"":"-P"}`;
  return {o, final, through, from, lines, rental, bundleDisc, delivery, charges, sales, adj, subtotal, tax, total:subtotal+tax, invDate, due,
    estimated: !final && !cn, overdue: isOverdue(o), cyc, cn, bill, prior, invNo};
}
function invoiceBody(d){
  const o = d.o, c = custById(o.customerId), l = locById(o.locationId), st = S.settings, e = esc;
  const dt = s => fmtDate(s), dtt = s => `${fmtDate(s)} ${s.slice(11,16)}`;
  const addr = (st.address||"").split("\n").filter(Boolean).map(e).join("<br>");
  return `<div class="inv">
  <header class="inv-head">
    <div>${st.logo?`<img class="inv-logo-img" src="${st.logo}" alt="${e(st.company)} logo">`:`<div class="inv-logo">BH<span>EQUIPMENT</span></div>`}<div class="inv-co"><b>${e(st.company)}</b>${addr?`<br>${addr}`:""}<br>${e(st.phone)}${st.email?` · ${e(st.email)}`:""}</div></div>
    <div class="inv-title"><h1>Invoice</h1><div class="inv-flag ${d.final||d.cn?"final":"open"}">${d.cn?`Cycle ${d.cn} · ${o.status==="Returned"?"equipment returned":"equipment on site"}`:d.final?(d.cyc?"Final · equipment returned":"Final · equipment returned"):d.overdue?"In progress · equipment overdue":"In progress · equipment on rent"}</div>
      <table class="inv-meta"><tr><th>Invoice #</th><td>${d.invNo}</td></tr><tr><th>Invoice date</th><td>${dt(d.invDate)}</td></tr><tr><th>Due</th><td>${e(st.terms||"Due on receipt")}${termsDays(st.terms)?` · ${dt(d.due)}`:""}</td></tr><tr><th>Order #</th><td>${o.no}</td></tr></table></div>
  </header>
  <section class="inv-parties">
    <div><h2>Bill to</h2><b>${e(c?.name||"—")}</b>${c?.position?`<br>${e(c.position)}`:""}${c?.company?`<br>${e(c.company)}`:""}${c?.phone?`<br>${e(c.phone)}`:""}${c?.email?`<br>${e(c.email)}`:""}</div>
    <div><h2>Job site</h2><b>${e(l?`${l.code} - ${l.name}`:"—")}</b>${l?.address?`<br>${e(l.address)}`:""}${locContact(l)?`<br>Site contact: ${e(locContact(l))}`:""}</div>
    <div><h2>Rental period</h2><b>Out:</b> ${dtt(o.out)}<br><b>${d.final?"Returned":"Due back"}:</b> ${dtt(d.final?(o.actualReturn||o.returnedAt||o.ret):o.ret)}${d.cyc?`<br><b>Billing:</b> ${e(cycleName(o))}<br><b>This invoice:</b> ${dt(d.from)} – ${dt(d.through)}`:`<br><b>Billed through:</b> ${dtt(d.through)}`}</div>
  </section>
  ${d.cn&&o.status!=="Returned"?`<p class="inv-note">Cycle ${d.cn} of the rental, ${dt(d.from)} to ${dt(d.through)}. The equipment is still on the job site${parseD(d.through) < parseD(cycEnd(o)) || o.status==="Rented Out" ? ` and the next cycle starts ${dt(d.through)}` : ""}. A final invoice is issued when it comes back.</p>`:""}
  ${d.final&&d.cyc&&d.prior.length?`<p class="inv-note">Already billed on cycle invoices: ${d.prior.map(b=>`INV-${o.no}-C${b.n} (${dt(b.from)} – ${dt(b.to)}, ${money(b.total)})`).join(", ")}. This final invoice covers the rest of the rental through the return, plus return charges.</p>`:""}
  ${d.estimated&&!d.cyc?`<p class="inv-note">${d.overdue?`This equipment is past its booked return of ${dt(o.ret)}. Rental is shown through today and will keep accruing until the equipment is checked in.`:`Rental is shown through the booked return date. A final invoice is issued when the equipment is checked in.`}</p>`:""}
  <table class="inv-lines">
    <thead><tr><th>Asset</th><th>Description</th><th class="r">Days</th><th>Billed as</th><th class="r">Amount</th></tr></thead>
    <tbody>
      ${d.final&&d.cyc&&!d.lines.length?`<tr><td></td><td colspan="4">All rental was billed on the cycle invoices.</td></tr>`:""}
      ${d.lines.map(x=>`<tr><td class="mono">#${e(x.id)}</td><td>${e(x.name)}${x.serial?`<div class="sub">S/N ${e(x.serial)}</div>`:""}<div class="sub">${x.rates}</div></td><td class="r">${x.days}</td><td>${e(x.label)}</td><td class="r">${money(x.total)}</td></tr>`).join("")}
      ${d.bundleDisc?`<tr><td></td><td>Bundle discount · ${e(o.bundle.name)} (${o.bundle.pct}%)</td><td></td><td></td><td class="r">-${money(d.bundleDisc)}</td></tr>`:""}
      ${d.delivery?`<tr><td></td><td>Delivery · ${e(o.delivery.name)}</td><td class="r">${o.delivery.trips}</td><td>${money(o.delivery.rate)} per trip</td><td class="r">${money(d.delivery)}</td></tr>`:""}
      ${d.sales.map(x=>`<tr><td class="mono">${e(x.sku)}</td><td>${e(x.name)}<div class="sub">Sold</div></td><td class="r">${x.qty}</td><td>${money(x.price)}${x.unit?" per "+e(x.unit):" each"}</td><td class="r">${money(x.total)}</td></tr>`).join("")}
      ${d.charges.map(x=>`<tr><td></td><td>${e(x.name)}</td><td class="r">${x.qty}</td><td>${money(x.rate)} each</td><td class="r">${money(x.total)}</td></tr>`).join("")}
      ${d.adj?`<tr><td></td><td>${d.adj<0?"Credit":"Adjustment"}</td><td></td><td></td><td class="r">${money(d.adj)}</td></tr>`:""}
    </tbody>
  </table>
  <div class="inv-bottom">
    <div class="inv-terms">
      ${o.description?`<p><b>Description:</b> ${e(o.description)}</p>`:""}
      ${d.final&&o.returnNotes?`<p><b>Return condition:</b> ${e(o.returnNotes)}</p>`:""}
      ${st.invoiceNote?`<p>${e(st.invoiceNote).replace(/\n/g,"<br>")}</p>`:""}
    </div>
    <table class="inv-totals">
      <tr><th>Subtotal</th><td>${money(d.subtotal)}</td></tr>
      ${d.tax?`<tr><th>Sales tax${+st.taxRate?` (${+st.taxRate}%)`:""}</th><td>${money(d.tax)}</td></tr>`:""}
      <tr class="grand"><th>${d.final||d.cn?"Amount due":"Estimated total"}</th><td>${money(d.total)}</td></tr>
    </table>
  </div>
  <footer class="inv-foot">${e(st.company)} · Order #${o.no}${d.bill?` · Billed ${dt(d.bill.at)}${d.bill.by?` by ${e(d.bill.by)}`:""}`:""} · Printed ${dtt(toLocalInput(now()))}${myName()?` by ${e(myName())}`:""}</footer>
</div>`;
}
const INVOICE_CSS = `
@page{size:letter;margin:.5in}
*{box-sizing:border-box}
body{margin:0;background:#fff;color:#16170f;font:12px/1.45 Arial,"Helvetica Neue",Helvetica,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.inv{max-width:7.5in;margin:0 auto;padding:.4in 0}
.inv-head{display:flex;justify-content:space-between;gap:24px;border-bottom:3px solid #4b5320;padding-bottom:16px}
.inv-logo{font:800 30px/1 Arial Black,Arial,sans-serif;letter-spacing:-.03em;display:flex;flex-direction:column}
.inv-logo span{font-size:8.5px;letter-spacing:.2em;font-weight:700;margin-top:2px}
.inv-logo-img{display:block;max-height:70px;max-width:3in;object-fit:contain}
.inv-co{margin-top:10px;font-size:11.5px;color:#3d4033}
.inv-title{text-align:right}
.inv-title h1{margin:0;font-size:30px;letter-spacing:.06em;text-transform:uppercase;color:#4b5320}
.inv-flag{display:inline-block;margin:6px 0 10px;padding:3px 9px;border-radius:3px;font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em}
.inv-flag.final{background:#4b5320;color:#fff}
.inv-flag.open{background:#fbf0d6;color:#7a5000;border:1px solid #e5c98a}
.inv-meta{margin-left:auto;border-collapse:collapse;font-size:11.5px}
.inv-meta th{text-align:right;color:#62665a;font-weight:400;padding:2px 10px 2px 0}
.inv-meta td{text-align:right;font-weight:700;padding:2px 0}
.inv-parties{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin:18px 0}
.inv-parties h2,.inv-terms h2{margin:0 0 4px;font-size:9.5px;letter-spacing:.12em;text-transform:uppercase;color:#62665a}
.inv-note{background:#fbf0d6;border-left:3px solid #c9a24a;padding:8px 10px;margin:0 0 14px;font-size:11.5px}
.inv-lines{width:100%;border-collapse:collapse}
.inv-lines th{background:#4b5320;color:#fff;text-align:left;font-size:10px;letter-spacing:.08em;text-transform:uppercase;padding:7px 8px}
.inv-lines td{padding:8px;border-bottom:1px solid #dfe1d4;vertical-align:top}
.inv-lines tbody tr:nth-child(even) td{background:#f6f7f0}
.r{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.mono{font-family:"Courier New",monospace;font-weight:700;white-space:nowrap}
.sub{color:#62665a;font-size:10.5px}
.inv-bottom{display:grid;grid-template-columns:1fr 2.6in;gap:24px;margin-top:16px;page-break-inside:avoid}
.inv-terms p{margin:0 0 8px}
.inv-totals{width:100%;border-collapse:collapse;align-self:start}
.inv-totals th{text-align:left;font-weight:400;padding:5px 0;color:#3d4033}
.inv-totals td{text-align:right;padding:5px 0;font-variant-numeric:tabular-nums}
.inv-totals .grand th,.inv-totals .grand td{border-top:2px solid #16170f;font-size:15px;font-weight:700;padding-top:8px;color:#16170f}
.inv-sign{display:grid;grid-template-columns:2fr 1.4fr 1fr;gap:24px;margin-top:44px;font-size:10.5px;color:#62665a;page-break-inside:avoid}
.inv-sign span{display:block;border-bottom:1px solid #16170f;height:26px;margin-bottom:4px}
.inv-foot{margin-top:28px;padding-top:8px;border-top:1px solid #dfe1d4;font-size:9.5px;color:#8a8d7f;text-align:center}
@media screen{body{background:#e9eadf}.inv{background:#fff;padding:.5in;margin:16px auto;box-shadow:0 2px 14px rgba(0,0,0,.15)}}
@media print{.inv{padding:0}}`;
function invoiceDoc(o, autoPrint, n){
  const d = invoiceData(o, {cycle:n});
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Invoice ${d.invNo} · ${esc(S.settings.company)}</title><style>${INVOICE_CSS}</style></head><body>${invoiceBody(d)}${autoPrint?`<script>window.addEventListener("load",function(){setTimeout(function(){window.print()},300)})<\/script>`:""}</body></html>`;
}
function openInvoice(no, n){
  const o = S.orders.find(x=>x.no===no); if (!o) return;
  n = o.cycle ? (+n||0) : 0;
  if (!["Rented Out","Returned"].includes(o.status)) { toast("Invoices print once the order is rented out or returned"); return; }
  if (o.cycle && !n && o.status!=="Returned"){ viewOrder(no); return; }
  const d = invoiceData(o, {cycle:n}), b = d.bill;
  const pending = n && !b;
  openModal(`<header><h2>Invoice ${d.invNo}</h2>${pending?`<span class="chip maint">Not billed yet</span>`:b?`<span class="chip available">Billed ${fmtDate(b.at)}</span>`:chipFor(d.final?"Returned":orderStatus(o))}<button class="x" data-act="close" aria-label="Close">×</button></header>
  <div class="mbody" style="padding:0"><iframe id="inv-frame" title="Invoice preview" class="inv-frame" srcdoc="${esc(invoiceDoc(o,false,n))}"></iframe></div>
  <footer><span class="total num">${d.final||n?"Amount due":"Estimated"} ${money(d.total)}</span>
    ${b && can("orders") ? `<button class="btn danger" data-act="cycle-void" data-no="${o.no}" data-n="${n}" style="margin-right:auto">Void</button>`:""}
    <button class="btn" data-act="view-order" data-no="${o.no}">Back to order</button>
    <button class="btn" data-act="inv-dl" data-no="${o.no}" data-n="${n}">${ico("download")} Download PDF</button>
    ${IN_VIEWER?"":`<button class="btn ${pending?"":"primary"}" data-act="inv-print" data-no="${o.no}" data-n="${n}">Print</button>`}
    ${pending?`<button class="btn primary" data-act="cycle-mark" data-no="${o.no}" data-n="${n}">Mark cycle ${n} billed</button>`:""}</footer>`, "invoice");
}
function printInvoice(no, n){
  const o = S.orders.find(x=>x.no===no); const f = document.createElement("iframe");
  f.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0"; document.body.appendChild(f);
  f.srcdoc = invoiceDoc(o, false, +n||0);
  f.onload = () => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch(e){ toast("Printing is blocked here. Use Download, then print the file."); } setTimeout(()=>f.remove(), 60000); };
}
/* ---------------- Invoice PDFs: a small built-in PDF writer, so downloads work offline with no add-ons ---------------- */
const PDF_CP = {"€":128,"‚":130,"„":132,"…":133,"‹":139,"‘":145,"’":146,"“":147,"”":148,"•":149,"–":150,"—":151,"™":153,"›":155};
function pdfEsc(s){ let out = ""; for (const ch of String(s ?? "")){ const c = ch.codePointAt(0); let b = PDF_CP[ch] ?? (c===0x2212 ? 45 : c===0x2192 ? 155 : c < 256 ? c : 63);
  out += b===40||b===41||b===92 ? "\\"+String.fromCharCode(b) : b<32||b>126 ? "\\"+b.toString(8).padStart(3,"0") : String.fromCharCode(b); } return out; }
let _pdfCtx = null;
function pdfW(s, size, bold){ _pdfCtx ||= document.createElement("canvas").getContext("2d"); _pdfCtx.font = `${bold?"bold ":""}100px Helvetica, Arial, "Liberation Sans", sans-serif`; return _pdfCtx.measureText(String(s ?? "")).width / 100 * size; }
const pdfRGB = h => { const n = parseInt(h.slice(1), 16); return [(n>>16)&255, (n>>8)&255, n&255].map(v=>(v/255).toFixed(3)).join(" "); };
class PdfDoc {
  constructor(){ this.W = 612; this.H = 792; this.pages = []; this.img = null; this.add(); }
  add(){ this.ops = []; this.pages.push(this.ops); }
  text(x, y, s, o={}){ const size = o.size||9, bold = !!o.bold, w = pdfW(s, size, bold), f = o.mono ? "F3" : bold ? "F2" : "F1";
    const xx = o.align==="right" ? x - w : o.align==="center" ? x - w/2 : x;
    this.ops.push(`BT /${f} ${size} Tf ${pdfRGB(o.color||"#16170f")} rg ${xx.toFixed(2)} ${(this.H-y).toFixed(2)} Td (${pdfEsc(s)}) Tj ET`); return w; }
  rect(x, y, w, h, fill){ this.ops.push(`${pdfRGB(fill)} rg ${x.toFixed(2)} ${(this.H-y-h).toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f`); }
  line(x1, y1, x2, y2, color="#dfe1d4", width=0.75){ this.ops.push(`${pdfRGB(color)} RG ${width} w ${x1.toFixed(2)} ${(this.H-y1).toFixed(2)} m ${x2.toFixed(2)} ${(this.H-y2).toFixed(2)} l S`); }
  image(x, y, w, h){ if (this.img) this.ops.push(`q ${w.toFixed(2)} 0 0 ${h.toFixed(2)} ${x.toFixed(2)} ${(this.H-y-h).toFixed(2)} cm /Im1 Do Q`); }
  wrap(s, size, bold, maxW){ const out = []; String(s ?? "").split("\n").forEach(par=>{ let cur = "";
    par.split(/\s+/).filter(Boolean).forEach(word=>{ const t = cur ? cur+" "+word : word; if (pdfW(t, size, bold) <= maxW || !cur) cur = t; else { out.push(cur); cur = word; } });
    out.push(cur); }); return out; }
  bytes(){
    const enc = s => { const a = new Uint8Array(s.length); for (let i=0;i<s.length;i++) a[i] = s.charCodeAt(i) & 255; return a; };
    const objs = [], font = n => `<< /Type /Font /Subtype /Type1 /BaseFont /${n} /Encoding /WinAnsiEncoding >>`;
    objs[3] = font("Helvetica"); objs[4] = font("Helvetica-Bold"); objs[5] = font("Courier-Bold");
    let next = 6; const imgId = this.img ? next++ : 0;
    if (this.img) objs[imgId] = [enc(`<< /Type /XObject /Subtype /Image /Width ${this.img.w} /Height ${this.img.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${this.img.bytes.length} >>\nstream\n`), this.img.bytes, enc("\nendstream")];
    const kids = this.pages.map(ops=>{ const c = ops.join("\n"), cid = next++, pid = next++;
      objs[cid] = `<< /Length ${c.length} >>\nstream\n${c}\nendstream`;
      objs[pid] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${this.W} ${this.H}] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >>${imgId?` /XObject << /Im1 ${imgId} 0 R >>`:""} >> /Contents ${cid} 0 R >>`; return pid; });
    objs[1] = `<< /Type /Catalog /Pages 2 0 R >>`; objs[2] = `<< /Type /Pages /Kids [${kids.map(k=>k+" 0 R").join(" ")}] /Count ${kids.length} >>`;
    const parts = [enc("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n")], offs = []; let len = parts[0].length;
    for (let i=1;i<next;i++){ offs[i] = len; const body = objs[i], chunks = [enc(`${i} 0 obj\n`), ...(Array.isArray(body) ? body : [enc(body)]), enc("\nendobj\n")]; chunks.forEach(c=>{ parts.push(c); len += c.length; }); }
    const xref = `xref\n0 ${next}\n0000000000 65535 f \n` + offs.slice(1).map(o=>String(o).padStart(10,"0")+" 00000 n \n").join("") + `trailer\n<< /Size ${next} /Root 1 0 R /Info << /Producer (BH Equipment Rentals) >> >>\nstartxref\n${len}\n%%EOF`;
    parts.push(enc(xref)); return new Blob(parts, {type:"application/pdf"});
  }
}
const _pdfLogo = {};
function pdfLogo(src){
  if (!src) return Promise.resolve(null);
  if (_pdfLogo[src]) return _pdfLogo[src];
  return _pdfLogo[src] = new Promise(res=>{ const im = new Image(); im.onload = () => { try {
      const k = Math.min(1, 600/im.naturalWidth), w = Math.max(1, Math.round(im.naturalWidth*k)), h = Math.max(1, Math.round(im.naturalHeight*k));
      const c = document.createElement("canvas"); c.width = w; c.height = h; const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0,0,w,h); x.drawImage(im,0,0,w,h);
      const b = atob(c.toDataURL("image/jpeg", 0.92).split(",")[1]), bytes = new Uint8Array(b.length); for (let i=0;i<b.length;i++) bytes[i] = b.charCodeAt(i);
      res({bytes, w, h}); } catch(e){ res(null); } }; im.onerror = () => res(null); im.src = src; });
}
async function invoicePdf(d){
  const o = d.o, c = custById(o.customerId), l = locById(o.locationId), st = S.settings;
  const P = new PdfDoc(), M = 42, R = P.W - M, CW = R - M, INK = "#16170f", MUTE = "#62665a", OLIVE = "#4b5320";
  const dt = s => fmtDate(s), dtt = s => `${fmtDate(s)} ${String(s).slice(11,16)}`;
  P.img = await pdfLogo(st.logo);
  /* header, left: logo + company */
  let y = M, ly;
  if (P.img){ const k = Math.min(52/P.img.h, 190/P.img.w); P.image(M, y, P.img.w*k, P.img.h*k); ly = y + P.img.h*k + 14; }
  else { P.text(M, y+24, "BH", {size:28, bold:true}); P.text(M, y+34, "E Q U I P M E N T", {size:6.5, bold:true}); ly = y + 50; }
  P.text(M, ly, st.company, {size:9.5, bold:true}); ly += 12;
  [...(st.address||"").split("\n").filter(Boolean), [st.phone, st.email].filter(Boolean).join(" · ")].filter(Boolean).forEach(t=>{ P.text(M, ly, t, {size:8.5, color:"#3d4033"}); ly += 11; });
  /* header, right: title, flag, meta */
  let ry = y + 22; P.text(R, ry, "INVOICE", {size:24, bold:true, color:OLIVE, align:"right"}); ry += 10;
  const flag = (d.cn ? `Cycle ${d.cn} · ${o.status==="Returned"?"equipment returned":"equipment on site"}` : d.final ? "Final · equipment returned" : d.overdue ? "In progress · equipment overdue" : "In progress · equipment on rent").toUpperCase();
  const fw = pdfW(flag, 7.5, true) + 14, solid = d.final || d.cn;
  P.rect(R-fw, ry, fw, 15, solid ? OLIVE : "#fbf0d6"); P.text(R-7, ry+10.5, flag, {size:7.5, bold:true, color:solid ? "#ffffff" : "#7a5000", align:"right"}); ry += 30;
  [["Invoice #", d.invNo], ["Invoice date", dt(d.invDate)], ["Due", `${st.terms||"Due on receipt"}${termsDays(st.terms)?` · ${dt(d.due)}`:""}`], ["Order #", String(o.no)]].forEach(([k,v])=>{
    const vw = P.text(R, ry, v, {size:9, bold:true, align:"right"}); P.text(R - vw - 10, ry, k, {size:9, color:MUTE, align:"right"}); ry += 13; });
  y = Math.max(ly, ry) + 4; P.line(M, y, R, y, OLIVE, 2.5); y += 20;
  /* parties */
  const colW = (CW - 28) / 3, cols = [
    ["Bill to", [[c?.name||"—", true], [c?.position], [c?.company], [c?.phone], [c?.email]]],
    ["Job site", [[l ? `${l.code ? l.code+" - " : ""}${l.name}` : "—", true], [l?.address], [locContact(l) ? `Site contact: ${locContact(l)}` : ""]]],
    ["Rental period", [[["Out: ", dtt(o.out)]], [[d.final?"Returned: ":"Due back: ", dtt(d.final?(o.actualReturn||o.returnedAt||o.ret):o.ret)]],
      ...(d.cyc ? [[["Billing: ", cycleName(o)]], [["This invoice: ", `${dt(d.from)} – ${dt(d.through)}`]]] : [[["Billed through: ", dtt(d.through)]]])]]];
  let py = y;
  cols.forEach(([h, rows], i)=>{ const x = M + i*(colW+14); let yy = y; P.text(x, yy, h.toUpperCase(), {size:7, bold:true, color:MUTE}); yy += 13;
    rows.forEach(([v, b])=>{ if (!v) return;
      if (Array.isArray(v)){ const lw = pdfW(v[0], 9, true); P.wrap(v[1], 9, false, colW - lw).forEach((t,j)=>{ if (!j) P.text(x, yy, v[0], {size:9, bold:true}); P.text(x + lw, yy, t, {size:9}); yy += 12; }); }
      else P.wrap(v, 9, !!b, colW).forEach(t=>{ P.text(x, yy, t, {size:9, bold:!!b}); yy += 12; }); });
    py = Math.max(py, yy); });
  y = py + 10;
  /* notes */
  const notes = [];
  if (d.cn && o.status!=="Returned") notes.push(`Cycle ${d.cn} of the rental, ${dt(d.from)} to ${dt(d.through)}. The equipment is still on the job site and the next cycle starts ${dt(d.through)}. A final invoice is issued when it comes back.`);
  if (d.final && d.cyc && d.prior.length) notes.push(`Already billed on cycle invoices: ${d.prior.map(b=>`INV-${o.no}-C${b.n} (${dt(b.from)} – ${dt(b.to)}, ${money(b.total)})`).join(", ")}. This final invoice covers the rest of the rental through the return, plus return charges.`);
  if (d.estimated && !d.cyc) notes.push(d.overdue ? `This equipment is past its booked return of ${dt(o.ret)}. Rental is shown through today and will keep accruing until the equipment is checked in.` : `Rental is shown through the booked return date. A final invoice is issued when the equipment is checked in.`);
  notes.forEach(n=>{ const ls = P.wrap(n, 8.5, false, CW - 22), h = ls.length*11 + 12; P.rect(M, y, CW, h, "#fbf0d6"); P.rect(M, y, 2.5, h, "#c9a24a"); ls.forEach((t,i)=>P.text(M+11, y+14+i*11, t, {size:8.5})); y += h + 12; });
  /* line items */
  const X = {asset:M+6, desc:M+64, days:M+CW-196, as:M+CW-186, amt:R-6}, descW = X.days - 30 - X.desc, BOTTOM = P.H - 64;
  const head = () => { P.rect(M, y, CW, 20, OLIVE); const o2 = {size:7, bold:true, color:"#ffffff"};
    P.text(X.asset, y+13.5, "ASSET", o2); P.text(X.desc, y+13.5, "DESCRIPTION", o2); P.text(X.days, y+13.5, "DAYS", {...o2, align:"right"}); P.text(X.as, y+13.5, "BILLED AS", o2); P.text(X.amt, y+13.5, "AMOUNT", {...o2, align:"right"}); y += 20; };
  const rows = [];
  if (d.final && d.cyc && !d.lines.length) rows.push({desc:"All rental was billed on the cycle invoices."});
  d.lines.forEach(x=>rows.push({asset:"#"+x.id, desc:x.name, subs:[x.serial?`S/N ${x.serial}`:"", x.rates].filter(Boolean), days:String(x.days), as:x.label, amt:money(x.total)}));
  if (d.bundleDisc) rows.push({desc:`Bundle discount · ${o.bundle.name} (${o.bundle.pct}%)`, amt:"-"+money(d.bundleDisc)});
  if (d.delivery) rows.push({desc:`Delivery · ${o.delivery.name}`, days:String(o.delivery.trips), as:`${money(o.delivery.rate)} per trip`, amt:money(d.delivery)});
  d.sales.forEach(x=>rows.push({asset:x.sku, desc:x.name, subs:["Sold"], days:String(x.qty), as:`${money(x.price)}${x.unit?" per "+x.unit:" each"}`, amt:money(x.total)}));
  d.charges.forEach(x=>rows.push({desc:x.name, days:String(x.qty), as:`${money(x.rate)} each`, amt:money(x.total)}));
  if (d.adj) rows.push({desc:d.adj<0?"Credit":"Adjustment", amt:money(d.adj)});
  head();
  rows.forEach((r, i)=>{ const dl = P.wrap(r.desc, 9, false, descW), sl = (r.subs||[]).flatMap(s=>P.wrap(s, 7.5, false, descW)), asl = P.wrap(r.as||"", 8.5, false, X.amt - 70 - X.as);
    const h = 10 + Math.max(dl.length*11.5 + sl.length*10, asl.length*11);
    if (y + h > BOTTOM){ P.add(); y = M; head(); }
    if (i % 2) P.rect(M, y, CW, h, "#f6f7f0");
    let ty = y + 15; if (r.asset) P.text(X.asset, ty, r.asset, {size:8.5, mono:true});
    dl.forEach(t=>{ P.text(X.desc, ty, t, {size:9}); ty += 11.5; }); sl.forEach(t=>{ P.text(X.desc, ty-1, t, {size:7.5, color:MUTE}); ty += 10; });
    if (r.days) P.text(X.days, y+15, r.days, {size:9, align:"right"}); asl.forEach((t,j)=>P.text(X.as, y+15+j*11, t, {size:8.5}));
    if (r.amt) P.text(X.amt, y+15, r.amt, {size:9, align:"right"});
    y += h; P.line(M, y, R, y); });
  /* totals + terms */
  const terms = [o.description?`Description: ${o.description}`:"", d.final&&o.returnNotes?`Return condition: ${o.returnNotes}`:"", st.invoiceNote||""].filter(Boolean);
  const tl = terms.flatMap(t=>[...P.wrap(t, 8.5, false, CW - 230), ""]), tot = [["Subtotal", money(d.subtotal)], ...(d.tax?[[`Sales tax${+st.taxRate?` (${+st.taxRate}%)`:""}`, money(d.tax)]]:[])];
  const need = Math.max(tl.length*11, tot.length*16 + 34) + 16;
  if (y + need > BOTTOM){ P.add(); y = M; }
  y += 16; const TX = R - 190;
  tl.forEach((t,i)=>{ if (t) P.text(M, y+10+i*11, t, {size:8.5, color:"#3d4033"}); });
  let ty = y + 10; tot.forEach(([k,v])=>{ P.text(TX, ty, k, {size:9, color:"#3d4033"}); P.text(R, ty, v, {size:9, align:"right"}); ty += 16; });
  P.line(TX, ty-6, R, ty-6, INK, 1.5); ty += 12;
  P.text(TX, ty, d.final||d.cn ? "Amount due" : "Estimated total", {size:12, bold:true}); P.text(R, ty, money(d.total), {size:12, bold:true, align:"right"});
  /* footer on every page */
  const foot = `${st.company} · Order #${o.no}${d.bill?` · Billed ${dt(d.bill.at)}${d.bill.by?` by ${d.bill.by}`:""}`:""} · Printed ${dtt(toLocalInput(now()))}${myName()?` by ${myName()}`:""}`;
  P.pages.forEach((ops, i)=>{ P.ops = ops; P.line(M, P.H-40, R, P.H-40); P.text(P.W/2, P.H-28, foot + (P.pages.length>1 ? ` · Page ${i+1} of ${P.pages.length}` : ""), {size:7.5, color:"#8a8d7f", align:"center"}); });
  return P.bytes();
}
async function downloadInvoice(no, n){ const o = S.orders.find(x=>x.no===no), d = invoiceData(o, {cycle:+n||0});
  let pdf; try { pdf = await invoicePdf(d); } catch(e){ console.error(e); toast("Couldn't build the PDF. Try again, or use Print."); return; }
  saveFile(`Invoice-${d.invNo}-${slug(custById(o.customerId)?.name||"customer")}.pdf`, pdf); }

/* ================= Phone layout: app bar, bottom tabs, quick actions ================= */
const isPhone = () => matchMedia("(max-width:760px)").matches;
function renderMobileChrome(){
  const bar = $("#m-appbar"), tabs = $("#m-tabbar"); if (!bar || !tabs) return;
  const s = S.settings, cur = route();
  bar.innerHTML = `<div class="m-title">${s.logo?`<img src="${s.logo}" alt="">`:""}<span>${esc(s.company||"BH Equipment")}</span></div>
    <button class="m-icon" data-act="m-search" aria-label="Search">${ico("search")}</button>
    <button class="m-icon" data-act="m-profile" aria-label="Your profile">${avatarHtml(me(),"sm")}</button>`;
  const on = k => k==="items" ? (ITEM_KEYS.includes(cur) || /^asset(-|$)/.test(cur)) : k==="more" ? !["dashboard","orders","items"].includes(cur) && !ITEM_KEYS.includes(cur) && !/^asset(-|$)/.test(cur) : cur===k;
  const tab = (k, label, icon) => `<a href="#${k}" class="m-tab" ${on(k)?'aria-current="page"':""}>${ico(icon)}<span>${label}</span></a>`;
  if (isSuper()){ const n = ui.req.cart.length;
    tabs.innerHTML = `<a href="#request" class="m-tab" ${cur==="request"?'aria-current="page"':""}>${ico("request")}<span>Request</span></a>` +
      `<button class="m-fab" data-act="${n?"rq-review":"goto"}" data-to="request" aria-label="Review request">${n?`<b style="font-size:18px">${n}</b>`:ico("plus",2.4)}</button>` +
      `<a href="#my-requests" class="m-tab" ${cur==="my-requests"?'aria-current="page"':""}>${ico("my-requests")}<span>My requests</span></a>` +
      `<button class="m-tab" data-act="m-more">${ico("menu")}<span>More</span></button>`;
    tabs.classList.add("super"); return; }
  tabs.classList.remove("super");
  tabs.innerHTML = tab("dashboard","Dashboard","dashboard") + tab("orders","Orders","orders") +
    `<button class="m-fab" data-act="m-fab" aria-label="New">${ico("plus",2.4)}</button>` + tab("items","Items","items") +
    `<button class="m-tab" data-act="m-more" ${on("more")?'aria-current="page"':""}>${ico("menu")}<span>More</span></button>`;
}
function sheet(title, body){
  openModal(`<header><h2>${esc(title)}</h2><button class="x" data-act="close" aria-label="Close">×</button></header><div class="mbody m-sheet">${body}</div>`, "sheet");
}
function sheetRow(act, icon, label, sub, extra=""){
  return `<button class="m-row" data-act="${act}" ${extra}><span class="m-row-ic">${ico(icon)}</span><span class="m-row-txt"><b>${label}</b>${sub?`<small>${sub}</small>`:""}</span><span class="m-chev">›</span></button>`;
}
function openFab(){
  const rows = [];
  if (can("orders")) rows.push(sheetRow("new-order","orders","New rental order","Book equipment out to a job"));
  if (can("checkout")) rows.push(sheetRow("co-new","checkout","Check out","Hand assets, stock or inventory to a crew"));
  if (can("orders") || can("checkout")) rows.push(sheetRow("m-quickin","arrowback","Quick check-in","Return a rental or a crew check-out"));
  if (can("shop")) rows.push(sheetRow("new-maint","maintenance","New service request","Report a problem or log service"));
  if (can("assets")) rows.push(sheetRow("new-item","items","Add asset","Put a new machine or tool in the fleet"));
  if (can("people")) rows.push(sheetRow("new-cust","customers","Add customer","Project manager or superintendent"));
  sheet("Create", rows.join("") || '<div class="empty">Your role can view everything but can\'t create records.</div>');
}
function openMore(){
  const u = me(), cur = route();
  if (isSuper()) { sheet("More", `<div class="m-group"><a class="m-row" href="#my-reports"><span class="m-row-ic">${ico("reports")}</span><span class="m-row-txt"><b>My reports</b><small>Your orders, on rent and costs by job</small></span><span class="m-chev">›</span></a></div><div class="m-group">${sheetRow("profile","customers","Your profile",`${esc(u?.name||"")} · ${esc(u?.role||"")}`)}${sheetRow("theme","moon","Switch light / dark","")}${sheetRow("logout","logout","Sign out","")}</div>`); return; }
  const link = (k, icon, label, sub) => `<a class="m-row ${cur===k?"on":""}" href="#${k}"><span class="m-row-ic">${ico(icon)}</span><span class="m-row-txt"><b>${label}</b>${sub?`<small>${sub}</small>`:""}</span><span class="m-chev">›</span></a>`;
  const open = S.custody.filter(c=>!c.inAt).length, shop = S.maintenance.filter(m=>m.status==="Open").length;
  sheet("More", `<div class="m-group">
    ${link("checkout","checkout","Check In / Out", open?`${open} out with crews`:"")}
    ${link("maintenance","maintenance","Maintenance", shop?`${shop} open work order${shop>1?"s":""}`:"")}
    ${link("asset-stock","stock","Asset Stock")}
    ${link("inventory","inventory","Inventory")}
    ${link("groups","groups","Groups")}
    ${link("locations","locations","Locations")}
    ${link("customers","customers","Customers")}
    ${link("reports","reports","Reports")}
    ${can("setup")?link("setup","setup","Setup"):""}
    ${can("users")?link("users","customers","Users"):""}
  </div>
  <div class="m-group">
    ${sheetRow("profile","customers","Your profile",`${esc(u?.name||"")} · ${esc(u?.role||"")}`)}
    ${sheetRow("theme","moon","Switch light / dark","")}
    ${sheetRow("logout","logout","Sign out","")}
  </div>`);
}
function openQuickCheckin(){
  const rows = [];
  S.orders.filter(o=>o.status==="Rented Out").sort((a,b)=>a.ret.localeCompare(b.ret)).forEach(o=>{
    const late = isOverdue(o);
    rows.push({q:orderSearchText(o), html:`<button class="m-row" data-act="checkin" data-no="${o.no}"><span class="m-row-ic ${late?"bad":""}">${ico("orders")}</span><span class="m-row-txt"><b>Order #${o.no} · ${esc(o.lines.map(l=>"#"+l.itemId).join(", "))}</b><small>${esc(locLabel(locById(o.locationId)))} · ${late?`<span style="color:var(--bad)">${Math.floor((now()-parseD(o.ret))/DAY)} days late</span>`:"due "+fmtDate(o.ret)}</small></span><span class="m-chev">›</span></button>`}); });
  S.custody.filter(c=>!c.inAt).sort((a,b)=>(a.dueAt||"9").localeCompare(b.dueAt||"9")).forEach(c=>{
    const late = custodyOverdue(c);
    rows.push({q:(c.ref+" "+c.holder+" "+custodyLabel(c)).toLowerCase(), html:`<button class="m-row" data-act="co-in" data-id="${c.id}"><span class="m-row-ic ${late?"bad":""}">${ico("checkout")}</span><span class="m-row-txt"><b>${esc(custodyLabel(c))}${c.qty>1?` × ${c.qty}`:""}</b><small>${esc(c.ref)} · ${esc(c.holder)}${c.dueAt?` · ${late?'<span style="color:var(--bad)">late</span>':"due "+fmtDate(c.dueAt)}`:""}</small></span><span class="m-chev">›</span></button>`}); });
  window._qci = rows;
  sheet("Quick check-in", `<div class="field"><input id="qci-q" type="search" placeholder="Search order, asset, person" aria-label="Search what to check in"></div><div id="qci-list" class="m-group">${rows.map(r=>r.html).join("") || '<div class="empty"><strong>Nothing is out</strong>Rentals and crew check-outs show up here.</div>'}</div>`);
}
function openSearchSheet(){
  sheet("Search", `<div class="field"><input id="m-q" type="search" placeholder="Orders, assets, jobs, people" aria-label="Search" autocomplete="off"></div><div id="m-results" class="m-results"><div class="small muted">Start typing to search everything.</div></div>`);
  setTimeout(()=>$("#m-q")?.focus(), 50);
}
function mobileDashboard(){
  const d = now(), active = S.orders.filter(o=>o.status==="Rented Out"), over = active.filter(isOverdue);
  const soon = active.filter(o=>!isOverdue(o) && parseD(o.ret)-d < 7*DAY), reserved = S.orders.filter(o=>o.status==="Booked");
  const queued = S.maintenance.filter(m=>m.status==="Open" && (m.stage==="Submitted"||m.stage==="Accepted")), working = S.maintenance.filter(m=>m.status==="Open" && !["Submitted","Accepted"].includes(m.stage));
  const coOpen = S.custody.filter(c=>!c.inAt), coLate = coOpen.filter(custodyOverdue);
  const circle = (act, icon, label, extra="") => `<button class="m-quick" data-act="${act}" ${extra}><span>${ico(icon)}</span>${label}</button>`;
  const row = (act, icon, label, n, alert, extra="") => `<button class="m-stat ${alert&&n?"alert":""}" data-act="${act}" ${extra}><span class="m-stat-ic">${ico(icon)}</span><span class="m-stat-lbl">${label}</span><span class="m-stat-n num ${alert&&n?"bad":""}">${n}</span><span class="m-chev">›</span></button>`;
  return `<div class="m-only m-dash">
    <div class="m-quicks">
      ${can("orders")||can("checkout")?circle("m-quickin","arrowback","Quick Check-In"):""}
      ${circle("goto","orders","Rented Out Orders",'data-to="orders" data-filter="out"')}
      ${circle("goto","orders","Returned Orders",'data-to="orders" data-otab="returned"')}
      ${can("checkout")?circle("co-new","checkout","Check Out"):circle("goto","checkout","Check-Outs",'data-to="checkout"')}
    </div>
    <h2 class="m-h">Needs attention</h2>
    ${row("goto","orders","Overdue returns",over.length,true,'data-to="orders" data-filter="overdue"')}
    ${row("goto","orders","Booked, ready to go out",S.orders.filter(isReady).length,true,'data-to="orders" data-filter="ready"')}
    ${row("goto","checkout","Crew check-outs past due",coLate.length,true,'data-to="checkout" data-cotab="overdue"')}
    ${row("goto","maintenance","Work orders to start",queued.length,true,'data-to="maintenance"')}
    ${row("goto","maintenance","Work orders in progress",working.length,false,'data-to="maintenance"')}
    <h2 class="m-h">Orders</h2>
    ${row("goto","orders","Active orders",active.length,false,'data-to="orders" data-otab="active"')}
    ${row("goto","orders","Due back this week",soon.length,false,'data-to="orders" data-filter="soon"')}
    ${row("goto","orders","Booked orders",reserved.length,false,'data-to="orders" data-filter="booked"')}
    ${row("goto","checkout","Out with crews",coOpen.length,false,'data-to="checkout" data-cotab="out"')}
    <h2 class="m-h">Fleet</h2>
  </div>`;
}
/* turn tables into cards on phones: label every cell with its column header */
/* resizable columns (desktop): drag the divider between column headers; double-click a divider to reset */
const colKey = t => KEY + "-cols-" + t.dataset.cols;
let colSuppress = false;
function initResizable(root){
  if (isPhone()) return;
  (root||document).querySelectorAll(".tablewrap > table").forEach(t=>{
    const ths = [...t.querySelectorAll("thead tr:first-child > th")]; if (ths.length < 3) return;
    t.classList.add("resizable");
    if (!t.dataset.cols) t.dataset.cols = route().replace(/^asset-.*/,"asset") + ":" + ths.map(th=>th.textContent.replace(/[▲▼]/g,"").trim()).join("|");
    let saved = null; try { saved = JSON.parse(localStorage.getItem(colKey(t))||"null"); } catch(e){}
    if (!Array.isArray(saved) || saved.length!==ths.length || saved.some(w=>!(w>0))) saved = null;
    if (saved){ t._colW = saved; layoutCols(t); }   // until someone drags, the table keeps its natural column widths
    ths.forEach((th,i)=>{ if (i < (ths[0].querySelector("input[type=checkbox]") ? 2 : 1)) return; const g = document.createElement("span"); g.className = "col-grip"; g.dataset.i = i - 1; g.title = "Drag to resize · double-click to reset"; th.appendChild(g); });
  });
}
function layoutCols(t){
  const w = t._colW, wrap = t.parentElement.clientWidth;
  let cg = t.querySelector("colgroup"); if (!cg){ cg = document.createElement("colgroup"); t.prepend(cg); }
  cg.innerHTML = w.map((x,i)=>`<col style="${i===w.length-1?"":`width:${x}px`}">`).join("");
  t.style.tableLayout = "fixed"; t.style.minWidth = "0"; t.style.width = Math.max(wrap, w.reduce((a,b)=>a+b,0)) + "px";
}
document.addEventListener("pointerdown", e=>{
  const g = e.target.closest(".col-grip"); if (!g) return;
  e.preventDefault(); e.stopPropagation();
  const t = g.closest("table"), i = +g.dataset.i, x0 = e.clientX;
  if (!t._colW){ t._colW = [...t.querySelectorAll("thead tr:first-child > th")].map(th=>Math.ceil(th.getBoundingClientRect().width)); layoutCols(t); }
  const w0 = t._colW[i];
  g.classList.add("drag"); document.body.classList.add("col-resizing"); g.setPointerCapture?.(e.pointerId);
  const move = ev=>{ t._colW[i] = Math.max(48, Math.round(w0 + ev.clientX - x0)); layoutCols(t); };
  const up = ()=>{ g.classList.remove("drag"); document.body.classList.remove("col-resizing"); window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up);
    try { localStorage.setItem(colKey(t), JSON.stringify(t._colW)); } catch(err){}
    colSuppress = true; setTimeout(()=>colSuppress = false, 0); };
  window.addEventListener("pointermove", move); window.addEventListener("pointerup", up);
}, true);
document.addEventListener("click", e=>{ if (colSuppress || e.target.closest(".col-grip")){ e.stopPropagation(); e.preventDefault(); colSuppress = false; } }, true);
document.addEventListener("dblclick", e=>{ const g = e.target.closest(".col-grip"); if (!g) return; const t = g.closest("table"); try { localStorage.removeItem(colKey(t)); } catch(err){} rerender(); }, true);
function cardify(root){
  (root||document).querySelectorAll(".tablewrap table").forEach(tb=>{
    const heads = [...tb.querySelectorAll("thead th")].map(th=>th.textContent.replace(/[▲▼]/g,"").trim());
    tb.querySelectorAll("tbody tr").forEach(tr=>[...tr.children].forEach((td,i)=>{ if (!td.hasAttribute("colspan")) td.dataset.label = heads[i]||""; }));
    tb.classList.add("cardable");
  });
}

/* ---------------- CSV / files ---------------- */
function csv(rows){ return rows.map(r=>r.map(v=>{ v = v==null?"":String(v); return /[",\n]/.test(v)?`"${v.replace(/"/g,'""')}"`:v; }).join(",")).join("\r\n"); }
async function saveFile(name, data, mime="application/octet-stream"){
  const blob = data instanceof Blob ? data : new Blob([data], {type:mime});
  if (window.claude && typeof window.claude.use === "function"){
    let dl = null; try { dl = await window.claude.use("downloads"); } catch(e) {}
    if (!dl){ toast("Downloads aren't available in this view. Open the HTML file directly to export."); return; }
    try { await dl.save({filename:name, data:blob}); toast(`Saved ${name}`); }
    catch(e){ const c = e && e.code; toast(c==="declined" ? "Download cancelled" : c==="rate_limited" ? "Finish the open download prompt first" : c==="rejected_extension" ? `This view can't save .${name.split(".").pop()} files. Open the HTML file directly to download it.` : "Downloads aren't available in this view. Open the HTML file directly to export."); }
    return;
  }
  try { const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 1000); toast(`Downloaded ${name}`); }
  catch(e){ toast("This browser blocked the download."); }
}
function download(name, text, type="text/csv"){ return saveFile(name, (type==="text/csv"?"\ufeff":"") + text, type+";charset=utf-8"); }
function dataUrlToBlob(u){ const [h, b64] = u.split(","); const mime = (h.match(/data:([^;]+)/)||[])[1] || "application/octet-stream"; const bin = atob(b64); const a = new Uint8Array(bin.length); for (let i=0;i<bin.length;i++) a[i] = bin.charCodeAt(i); return new Blob([a], {type:mime}); }
function exportOrders(list){ list = list || S.orders;
  download(`orders-${toLocalInput(now()).slice(0,10)}.csv`, csv([["Order #","Customer","Location","Items","Rent out","Return date","Status","Description","Pickup","Rental","Delivery","Charges","Adjustment","Total"]].concat(list.map(o=>{const t=orderTotals(o); return [o.no,custById(o.customerId)?.name,locLabel(locById(o.locationId)),o.lines.map(l=>`#${l.itemId} ${itemById(l.itemId)?.name||""}`).join("; "),o.out.replace("T"," "),orderRentalEnd(o).replace("T"," "),orderStatus(o),o.description,o.pickup,t.rental.toFixed(2),t.delivery.toFixed(2),t.charges.toFixed(2),(+o.adjustment||0).toFixed(2),t.total.toFixed(2)];})))); }
function exportItems(){ download("assets.csv", csv([["Asset #","Description","Category","Location","Spot / bay","Day","Week","Month","Status","Serial","Utilization 90d","Notes"]].concat(S.items.map(i=>[i.id,i.name,i.category,locLabel(homeLoc(i)),i.home||"",i.day,i.week,i.month,itemStatus(i),i.serial,Math.round(utilization(i.id)*100)+"%",i.notes])))); }
async function fileSave(){
  const text = JSON.stringify(S,null,1);
  if (window.showSaveFilePicker){ try{ fileHandle = await window.showSaveFilePicker({suggestedName:"bh-equipment-data.json",types:[{description:"Rental data",accept:{"application/json":[".json"]}}]}); await writeFile(); toast(`Saving to ${fileHandle.name}`); rerender(); return; } catch(e){ if (e.name==="AbortError") return; } }
  download("bh-equipment-data.json", text, "application/json");
}
async function writeFile(){ const w = await fileHandle.createWritable(); await w.write(JSON.stringify(S,null,1)); await w.close(); }
async function fileOpen(){
  if (window.showOpenFilePicker){ try{ const [h] = await window.showOpenFilePicker({types:[{description:"Rental data",accept:{"application/json":[".json"]}}]}); const f = await h.getFile(); loadData(await f.text()); fileHandle = h; try{ await h.requestPermission({mode:"readwrite"}); }catch(e){} rerender(); return; } catch(e){ if (e.name==="AbortError") return; } }
  $("#importfile").click();
}
function loadData(text){ try{ const d = JSON.parse(text); if (!d.items||!d.orders) throw 0; S = migrate(d); save(); toast("Data file loaded"); render(); } catch(e){ toast("That file isn't a rental data file. Pick a .json saved from this app."); } }
$("#importfile").addEventListener("change", async e => { const f = e.target.files[0]; if (f) loadData(await f.text()); e.target.value=""; });

/* ---------------- Events ---------------- */
document.addEventListener("click", e => {
  const nav = e.target.closest('a[href^="#"]');
  if (nav){
    const to = nav.getAttribute("href").slice(1);
    if (to==="items") Object.assign(ui.items, {tab:"all", q:"", cat:"", loc:""});
    if (to==="orders") Object.assign(ui.orders, {filter:"", status:"", q:""});
    $("#flyout")?.remove(); if (nav.closest("#modal-root")) closeModal();
    if (location.hash.slice(1)===to || (!location.hash && to==="dashboard")){ e.preventDefault(); render(); return; }
  }
  const el = e.target.closest("[data-act]");
  if (!e.target.closest(".dd")) document.querySelectorAll(".dd-menu").forEach(m=>m.hidden=true);
  if (!e.target.closest(".gsearch")) $("#gresults").hidden = true;
  if (!e.target.closest("#pmenu") && !e.target.closest("#avatar-btn")) $("#pmenu")?.remove();
  if (!e.target.closest("#notif") && !e.target.closest("#bell-btn")) $("#notif")?.remove();
  if (!el) return;
  const a = el.dataset.act, ds = el.dataset;
  if (a==="file-open" && !me()) { fileOpen(); return; }   // allowed from the sign-in screen so a team data file can be loaded before signing in
  if (ACT_PERM[a] && !can(ACT_PERM[a])) { toast(`Your role (${me()?.role||"signed out"}) can't do that`); return; }
  if (a==="scrim"){ if (e.target===el) { if (modalHasForm()) { nudgeModal("Your changes are still here. Use Cancel or × to close without saving."); return; } closeModal(); draft=null; ci=null; } return; }
  if (["cmb-pick","cmb-mode","pick","ci-maint","sel","selall","co-pick","co-sel","co-selall","coin-maint"].includes(a)) { /* handled on change */ return; }
  if (a.startsWith("sr-") && srAction(a, ds)) return;
  if (a.startsWith("rq-") && reqAction(a, ds)) return;
  const no = ds.no ? +ds.no : null;
  switch(a){
    case "profile-menu": toggleProfileMenu(); break;
    case "bell": toggleNotif(); break;
    case "profile": $("#pmenu")?.remove(); closeModal(); profileForm(!!ds.pw); break;
    case "profile-save": saveProfile(); break;
    case "logout": $("#pmenu")?.remove(); closeModal(); setSession(null); render(); toast("Signed out"); break;
    case "user-new": userForm(null); break;
    case "user-edit": userForm(ds.id); break;
    case "user-save": saveUser(); break;
    case "user-del": { const id = ds.id, u = S.users.find(x=>x.id===id); if (u.role==="Admin" && !S.users.some(x=>x.id!==id && x.role==="Admin" && x.active)) { toast("Keep at least one active Admin"); break; }
      confirmBox(`Delete ${u.name}?`,"They won't be able to sign in. Their name stays on past orders and check-outs. To pause access instead, set Access to <b>Turned off</b>.","Delete user",()=>{ S.users = S.users.filter(x=>x.id!==id); save(); rerender(); toast("User deleted"); }); break; }
    case "close": pendingPhoto=undefined; closeModal(); draft=null; ci=null; codraft=null; coin=null; mdraft=null; break;
    case "cotab": ui.co.tab = ds.v; ui.co.sel.clear(); rerender(); break;
    case "co-new": if (ds.in) openCustodyCheckin(ui.co.sel.size?[...ui.co.sel]:S.custody.filter(c=>!c.inAt).map(c=>c.id)); else openCheckoutForm(); break;
    case "co-item": openCheckoutForm({assets:[ds.id]}); break;
    case "co-save": saveCheckout(); break;
    case "co-add": { const inp = $("#qa-"+ds.id), q = +(inp?.value||1); if (!(q>0)){ toast("Enter a quantity above 0"); break; }
      const list = ds.kind==="inv" ? codraft.inv : codraft.stock, key = ds.kind==="inv" ? "invId" : "stockId"; const ex = list.find(l=>l[key]===ds.id);
      if (ex) ex.qty = +(+ex.qty + q).toFixed(2); else list.push({[key]:ds.id, qty:q}); drawCheckoutForm(); refocus("co-pq"); break; }
    case "co-line-del": if (ds.kind==="asset") codraft.assets = codraft.assets.filter(x=>x!==ds.id); else (ds.kind==="inv"?codraft.inv:codraft.stock).splice(+ds.i,1); drawCheckoutForm(); break;
    case "inv-co": openCheckoutForm({inv:[{invId:ds.id,qty:1}], type:"inv"}); break;
    case "co-in": openCustodyCheckin([ds.id]); break;
    case "co-bulkin": openCustodyCheckin([...ui.co.sel]); break;
    case "coin-save": saveCustodyCheckin(); break;
    case "co-undo": { const id = ds.id, c = S.custody.find(x=>x.id===id); confirmBox(`Undo check-in of ${esc(custodyLabel(c))}?`, `It goes back to checked out to <b>${esc(c.holder)}</b>${c.kind==="stock"&&c.missingQty?`, and the ${c.missingQty} marked missing are added back to the total`:""}${c.kind==="inv"&&c.returnedQty?`, and the ${c.returnedQty} returned come back off the shelf`:""}.`, "Undo check-in", ()=>undoCheckin(id), false); break; }
    case "export-co": download("check-outs.csv", csv([["Ref","Type","Item","Qty","Unit","Returned qty","Used / missing","Assigned to","Location","Checked out","Due back","Returned","Out notes","In notes"]].concat(S.custody.map(c=>{ const unit = c.kind==="inv"?S.inventory.find(x=>x.id===c.invId)?.unit:c.kind==="stock"?stockById(c.stockId)?.unit:""; return [c.ref,{asset:"Asset",stock:"Asset stock",inv:"Inventory"}[c.kind],custodyLabel(c),c.qty,unit||"",c.returnedQty??"",c.kind==="inv"?(c.usedQty??""):c.kind==="stock"?(c.missingQty??""):"",c.holder,locLabel(locById(c.locationId)),c.outAt.replace("T"," "),(c.dueAt||"").replace("T"," "),(c.inAt||"").replace("T"," "),c.outNotes,c.inNotes]; })))); break;
    case "stock-new": stockForm(null); break;
    case "stock-edit": stockForm(ds.id); break;
    case "stock-save": saveStock(); break;
    case "stock-co": openCheckoutForm({stock:[{stockId:ds.id,qty:1}], type:"stock"}); break;
    case "stock-del": { const id=ds.id; if (stockOut(id)){ toast("Check the stock back in before deleting it"); break; } confirmBox("Delete this stock item?","Its check-out history keeps the record but loses the name.","Delete",()=>{ S.stock=S.stock.filter(s=>s.id!==id); save(); rerender(); }); break; }
    case "inv-new": invForm(null); break;
    case "inv-edit": invForm(ds.id); break;
    case "inv-save": saveInv(); break;
    case "inv-adj": invAdjust(ds.id); break;
    case "inv-adj-save": saveAdjust(); break;
    case "inv-del": { const id=ds.id; confirmBox("Delete this inventory item?","Work orders that used it keep their cost but lose the part name.","Delete",()=>{ S.inventory=S.inventory.filter(x=>x.id!==id); save(); rerender(); }); break; }
    case "export-inv": download("inventory.csv", csv([["Item","Group","SKU","On hand","Unit","Reorder at","Status","Unit cost","Charge price","Vendor","Location"]].concat(S.inventory.map(i=>[i.name,i.category||"",i.sku,i.qty,i.unit,i.reorder,invStatus(i),i.unitCost,i.price,i.vendor,keptAt(i)])))); break;
    case "m-addpart": { const inv = S.inventory.find(x=>!x.retired); if (!inv){ toast("All inventory is retired"); break; } mdraft.parts.push({invId:inv.id,qty:1,cost:+inv.unitCost||0}); $("#m-parts").innerHTML = partsEditor(); break; }
    case "m-delpart": mdraft.parts.splice(+ds.i,1); $("#m-parts").innerHTML = partsEditor(); break;
    case "group-new": groupForm(null, ds.v||""); break;
    case "group-edit": groupForm(ds.v); break;
    case "group-save": saveGroup(); break;
    case "group-combine": combineForm(ds.v); break;
    case "group-combine-save": saveCombine(); break;
    case "group-ungroup": { const g = groupOf(ds.v), par = g?.parent; if (g){ delete g.parent; save(); closeModal(); rerender(); toast(`${g.name} is a top-level group again${par?` (no longer under ${par})`:""}`); } break; }
    case "group-del": { const nm=ds.v; S.groups=S.groups.filter(g=>g.name!==nm); save(); closeModal(); rerender(); toast(`Group ${nm} deleted`); break; }
    case "items-all": Object.assign(ui.items, {tab:"all", q:"", cat:"", loc:""}); rerender(); break;
    case "group-assets": ui.items.loc = ""; ui.items.cat = ds.v; ui.items.exact = ds.exact ? ds.v : ""; ui.items.tab = "all"; go("items"); break;
    case "group-open": ui.groups.open = ds.v; if (route()==="groups") { rerender(); window.scrollTo(0,0); } else go("groups"); break;
    case "group-close": ui.groups.open = null; rerender(); break;
    case "req-pickup2": { const o = S.orders.find(x=>x.no===no); o.pickup="BHESC Pickup"; (o.log||=[]).push({by:myName(),at:toLocalInput(now()),text:"Pickup status set to BHESC Pickup"}); save(); rerender(); toast(`#${no} set to BHESC Pickup`); break; }
    case "goto": $("#notif")?.remove(); closeModal(); if (ds.mtab) ui.maint.tab = ds.mtab; if (ds.filter){ ui.orders.filter = ds.filter; ui.orders.tab = "active"; } if (ds.otab){ ui.orders.tab = ds.otab==="reserved" ? "active" : ds.otab; ui.orders.filter = ds.otab==="reserved" ? "booked" : (ds.filter||""); } if (ds.cotab) ui.co.tab = ds.cotab; go(ds.to); break;
    case "theme": { const r = document.documentElement; const dark = r.dataset.theme ? r.dataset.theme==="dark" : matchMedia("(prefers-color-scheme: dark)").matches; r.dataset.theme = dark?"light":"dark"; try{localStorage.setItem(KEY+"-theme",r.dataset.theme);}catch(_){} break; }
    case "new-order": openOrderForm(null); break;
    case "rent-item": if (itemById(ds.id)?.shopOnly){ toast(`#${ds.id} is shop use only. Turn on "Available to rent" to rent it.`); break; } draft = newOrderDraft(ds.id); drawOrderForm(); break;
    case "view-order": viewOrder(no); break;
    case "edit-order": openOrderForm(S.orders.find(o=>o.no===no)); break;
    case "dup-order": { const o = S.orders.find(x=>x.no===no); const d = newOrderDraft(); Object.assign(d,{customerId:o.customerId,locationId:o.locationId,lines:o.lines.map(l=>lineFrom(l.itemId)).filter(l=>itemById(l.itemId)),description:o.description,cycle:o.cycle,delivery:{...o.delivery},out:o.out,ret:o.ret}); draft=d; drawOrderForm(); break; }
    case "save-order": saveOrder(); break;
    case "unpick": draft.lines = draft.lines.filter(l=>l.itemId!==ds.id); drawOrderForm(); break;
    case "dw-edit": ui.dash.edit = !ui.dash.edit; rerender(); break;
    case "dw-reset": { const u = me(); if (u){ delete u.dash; save(); } rerender(); toast("Dashboard reset to the default layout"); break; }
    case "dw-remove": dashEdit(l=>l.splice(+ds.i,1)); break;
    case "dw-size": { const id = dashLayout()[+ds.i].id; ui.dash.snap = id; dashEdit(l=>{ const w = l[+ds.i]; w.size = ds.v; w.x = Math.min(w.x, DCOLS - dwSpan(w).w); dwResolve(l, +ds.i); }); setTimeout(()=>{ if (ui.dash.snap===id) ui.dash.snap = null; }, 700); break; }
    case "dw-add": { ui.dash.snap = ds.v; dashEdit(l=>{ const w = {id:ds.v, size: isStatW(ds.v) ? "s" : "m"}; Object.assign(w, dwSpot(l, w)); l.push(w); }); setTimeout(()=>{ if (ui.dash.snap===ds.v) ui.dash.snap = null; }, 700); break; }
    case "add-sale": { const it = S.inventory.find(i=>!i.retired && canSell(i)); if (!it) break; (draft.sales ||= []).push({invId:it.id, qty:1, price:+it.price||0}); drawOrderForm(); break; }
    case "del-sale": draft.sales.splice(+ds.i,1); drawOrderForm(); break;
    case "add-charge": { const c = S.settings.charges[0]; draft.charges.push({name:c.name,qty:1,rate:c.rate}); drawOrderForm(); break; }
    case "del-charge": (ci?ci.charges:draft.charges).splice(+ds.i,1); ci?drawCheckin():drawOrderForm(); break;
    case "mark-out": { const o = S.orders.find(x=>x.no===no); const bad = o.lines.filter(l=>openMaint(l.itemId)||activeOrderFor(l.itemId)||openCustody(l.itemId)); if (bad.length){ toast(`#${bad[0].itemId} isn't back in the yard yet`); break; } const nowStr2 = toLocalInput(now()), early = nowStr2 < o.out; if (early){ (o.log||=[]).push({by:myName(),at:nowStr2,text:`Went out early. Rental start moved from ${fmtDate(o.out)} ${o.out.slice(11)} to now`}); o.out = nowStr2; } o.status="Rented Out"; (o.log||=[]).push({by:myName(),at:nowStr2,text:"Rented out"}); syncSales(o); save(); viewOrder(no); rerender(); toast(`Order #${no} rented out${early?" (billing starts now)":""}`); break; }
    case "req-pickup": { const o = S.orders.find(x=>x.no===no); o.pickup="BHESC Pickup"; (o.log||=[]).push({by:myName(),at:toLocalInput(now()),text:"Pickup status set to BHESC Pickup"}); save(); viewOrder(no); rerender(); toast("Set to BHESC Pickup"); break; }
    case "cancel-order": confirmBox(`Cancel order #${no}?`,"The equipment on it becomes available again. The order stays in history as Cancelled.","Cancel order",()=>{ const o=S.orders.find(x=>x.no===no); o.status="Cancelled"; (o.log||=[]).push({by:myName(),at:toLocalInput(now()),text:"Cancelled"}); syncSales(o, "order cancelled"); save(); rerender(); toast(`Order #${no} cancelled`); }); break;
    case "confirm-ok": { const cb = confirmCb; confirmCb=null; closeModal(); cb && cb(); break; }
    case "checkin": openCheckin(no); break;
    case "ci-quick": { const c = S.settings.charges.find(x=>x.name===ds.v); ci.charges.push({name:c.name,qty:1,rate:c.rate}); drawCheckin(); break; }
    case "ci-save": saveCheckin(!!ds.print); break;
    case "inv-open": openInvoice(no, ds.n); break;
    case "inv-print": printInvoice(no, ds.n); break;
    case "inv-dl": downloadInvoice(no, ds.n); break;
    case "cycle-mark": markCycleBilled(no, +ds.n); break;
    case "cycle-void": voidCycleBill(no, +ds.n); break;
    case "otab": ui.orders.tab = ds.v==="reserved" ? "active" : ds.v; ui.orders.filter = ds.v==="reserved" ? "booked" : ""; rerender(); break;
    case "ofilter": ui.orders.filter = ds.v; ui.orders.tab = "active"; rerender(); break;
    case "sort": ui.orders.dir = ui.orders.sort===ds.k ? -ui.orders.dir : 1; ui.orders.sort = ds.k; rerender(); break;
    case "toggle-dd": { const m = $("#dd-"+ds.dd); m.hidden = !m.hidden; break; }
    case "bulk": { const sel = [...ui.orders.sel].map(n=>S.orders.find(o=>o.no===n)).filter(Boolean);
      if (ds.v==="pickup"){ sel.filter(o=>o.status==="Rented Out").forEach(o=>{o.pickup="BHESC Pickup"; (o.log||=[]).push({by:myName(),at:toLocalInput(now()),text:"Pickup status set to BHESC Pickup"});}); save(); toast(`${sel.length} order(s) set to BHESC Pickup`); }
      if (ds.v==="checkin"){ const at = toLocalInput(now()); let n=0; sel.filter(o=>o.status==="Rented Out").forEach(o=>{o.status="Returned"; o.returnedAt=at; o.actualReturn=at; if (o.pickup!=="BHESC Pickup") o.pickup="Customer Returned"; (o.log||=[]).push({by:myName(),at,text:"Checked in (bulk)"}); n++;}); save(); toast(`${n} order(s) checked in, billed through today`); }
      if (ds.v==="export") exportOrders(sel);
      ui.orders.sel.clear(); rerender(); break; }
    case "export-orders": exportOrders(); break;
    case "export-items": exportItems(); break;
    case "rpt-csv": exportReportCsv(ds.i); break;
    case "rpt-xlsx": exportReportXlsx(); break;
    case "rpt-all": exportAllXlsx(); break;
    case "rloc-pick": ui.reports.loc = ds.id; rerender(); break;
    case "itab": ui.items.tab = ds.v; rerender(); break;
    case "mtab": ui.maint.tab = ds.v; rerender(); break;
    case "rtab": ui.reports.tab = ds.v; rerender(); break;
    case "view-item": openAsset(ds.id); break;
    case "atab": ui.asset.tab = ds.v; rerender(); break;
    case "abtab": ui.asset.bottom = ds.v; rerender(); break;
    case "a-go": openAsset(ds.id); break;
    case "a-retire": retireForm("asset", ds.id); break;
    case "a-restore": restoreRetired("asset", ds.id); break;
    case "inv-retire": retireForm("inv", ds.id); break;
    case "inv-restore": restoreRetired("inv", ds.id); break;
    case "invtab": ui.inv.tab = ds.v; rerender(); break;
    case "retire-save": saveRetire(); break;
    case "m-fab": openFab(); break;
    case "m-more": case "m-profile": openMore(); break;
    case "m-search": openSearchSheet(); break;
    case "m-quickin": openQuickCheckin(); break;
    case "logo-up": pickFiles("image/png,image/jpeg,image/svg+xml,image/webp,image/gif", false, async f=>{ if (!f[0]) return; try { S.settings.logo = await resizeLogo(f[0]); save(); rerender(); toast("Logo updated"); } catch(e){ toast("Couldn't read that image. Use a PNG, JPG or SVG."); } }); break;
    case "logo-del": delete S.settings.logo; save(); rerender(); toast("Logo removed"); break;
    case "m-addfile": pickFiles("", true, f=>addDocs(f)); break;
    case "m-file-del": { const i = +ds.i; const f = mdraft.files[i]; mdraft.files.splice(i,1); $("#m-files").innerHTML = docsEditor(); toast(`${f.name} removed. Save the work order to confirm.`); break; }
    case "m-file-open": { const f = ds.mid ? S.maintenance.find(x=>x.id===ds.mid).files[+ds.i] : mdraft.files[+ds.i]; downloadData(f.name, f.data); break; }
    case "m-docs": viewDocs(ds.id); break;
    case "x-photo-pick": pickFiles("image/*", false, async f=>{ if (!f[0]) return; try { pendingPhoto = await resizeImage(f[0], 900, 0.8); const w=$("#fp-wrap"); if (w) w.innerHTML = photoPreview(pendingPhoto); } catch(e){ toast("Couldn't read that image. Use a JPG or PNG."); } }); break;
    case "x-photo-clear": { pendingPhoto = null; const w=$("#fp-wrap"); if (w) w.innerHTML = photoPreview(null); break; }
    case "x-photo-view": viewXPhoto(ds.kind, ds.id); break;
    case "dup-item": closeModal(); itemForm(null, ds.id); break;
    case "a-photo-add": pickFiles("image/*", true, f=>addPhotos(ds.id, f)); break;
    case "a-photo-view": viewPhoto(ds.id, +ds.i); break;
    case "a-photo-primary": { const it = itemById(ds.id); it.photos.forEach((p,i)=>p.primary = i===+ds.i); save(); rerender(); toast("Main photo set"); break; }
    case "a-photo-del": { const id=ds.id, i=+ds.i; confirmBox("Delete this photo?","It's removed from the asset for everyone.","Delete photo",()=>{ const it=itemById(id); const was=it.photos[i].primary; it.photos.splice(i,1); if (was && it.photos[0]) it.photos[0].primary=true; itemLog(it,"Photo deleted"); save(); rerender(); }); break; }
    case "a-comment": { const tx = $("#a-cmt").value.trim(); if (!tx){ toast("Type a comment first"); break; } const it = itemById(ds.id); (it.comments ||= []).push({id:uid("c"),at:toLocalInput(now()),by:myName(),text:tx}); save(); rerender(); toast("Comment posted"); break; }
    case "a-comment-del": { const it = itemById(ds.id), c = it.comments.find(x=>x.id===ds.c); if (c.by!==myName() && me().role!=="Admin"){ toast("Only the author or an Admin can delete a comment"); break; } it.comments = it.comments.filter(x=>x.id!==ds.c); save(); rerender(); break; }
    case "a-file-add": pickFiles("", true, f=>addFiles(ds.id, f)); break;
    case "a-file-open": openFile(ds.id, +ds.i); break;
    case "a-file-dl": { const f = itemById(ds.id).files[+ds.i]; downloadData(f.name, f.data); break; }
    case "a-file-del": { const id=ds.id, i=+ds.i, it=itemById(id); confirmBox(`Delete ${esc(it.files[i].name)}?`,"The attachment is removed for everyone.","Delete file",()=>{ it.files.splice(i,1); save(); rerender(); }); break; }
    case "a-hours": hoursForm(ds.id); break;
    case "a-hours-save": { const it=itemById(ds.id), v=$("#hr-val").value; if (v===""||isNaN(+v)){ toast("Enter the meter reading"); break; } const old=it.hours; it.hours=+v; itemLog(it,`Hours ${old??"—"} → ${it.hours}${$("#hr-note").value?" ("+$("#hr-note").value+")":""}`); save(); closeModal(); rerender(); toast("Hours updated"); break; }
    case "u-photo": { const id = ds.id; if (id!==SESSION_ID && !can("users")){ toast("You can only change your own photo"); break; } pickFiles("image/*", false, async f=>{ if (!f[0]) return; try { const u=S.users.find(x=>x.id===id); u.photo = await resizeImage(f[0], 256, 0.85); save(); renderRail(); if ($("#profileform")) profileForm(); else if ($("#userform")) userForm(id); rerender(); toast("Photo updated"); } catch(e){ toast("Couldn't read that image. Use a JPG or PNG."); } }); break; }
    case "u-photo-del": { const id = ds.id; if (id!==SESSION_ID && !can("users")) break; const u=S.users.find(x=>x.id===id); delete u.photo; save(); renderRail(); if ($("#profileform")) profileForm(); else userForm(id); rerender(); break; }
    case "new-item": itemForm(null); break;
    case "edit-item": itemForm(ds.id); break;
    case "save-item": saveItem(); break;
    case "del-item": { const id = ds.id; if (S.orders.some(o=>ACTIVE.includes(o.status)&&o.lines.some(l=>l.itemId===id))){ toast(`#${id} is on an open order. Close that order first, or mark the asset Retired.`); break; }
      confirmBox(`Delete asset #${id}?`,"Past orders keep the line but lose the description. Marking it <b>Retired</b> keeps its history instead.","Delete asset",()=>{ S.items = S.items.filter(i=>i.id!==id); save(); if (!VIEWS[route()]) go("items"); else rerender(); toast(`#${id} deleted`); }); break; }
    case "new-maint": maintForm(null, ds.id); break;
    case "edit-maint": maintForm(ds.id); break;
    case "save-maint": saveMaint(); break;
    case "view-sr": viewSR(ds.id); break;
    case "rqtab": ui.req.tab = ds.v; ui.req.cat = ""; rerender(); break;
    case "mytab": ui.myrep.tab = ds.v; rerender(); break;
    case "avtab": ui.av.tab = ds.v; ui.av.cat = ""; rerender(); break;
    case "av-csv": { const D = availData(), sh = D.sheets[{assets:0,stock:1,inv:2}[ui.av.tab]||0]; download(`Available-${slug(sh.name)}-${stamp()}.csv`, sheetCsv(sh)); break; }
    case "av-xlsx": saveFile(`Available-equipment-${stamp()}.xlsx`, buildXlsx(availData().sheets)); break;
    case "my-order": myOrderView(ds.no); break;
    case "my-csv": { const D = myRepData(), sh = D.sheets[{onrent:0,history:1,jobs:2}[ui.myrep.tab]||0]; download(`My-${slug(sh.name)}-${stamp()}.csv`, sheetCsv(sh)); break; }
    case "my-xlsx": saveFile(`My-equipment-report-${stamp()}.xlsx`, buildXlsx(myRepData().sheets)); break;
    case "rqatab": ui.rqa.tab = ds.v; rerender(); break;
    case "inv-clear": Object.assign(ui.inv, {q:"", cat:""}); rerender(); break;
    case "complete-maint": viewSR(ds.id); if (srById(ds.id).stage==="QC inspection"){ srUI.qc = true; drawSR(); } break;
    case "complete-maint-old": { const m = S.maintenance.find(x=>x.id===ds.id); m.status="Completed"; m.stage="Done"; m.doneAt=toLocalInput(now()).slice(0,10); save(); rerender(); toast(`#${m.itemId} is back in service`); break; }
    case "del-maint": confirmBox("Delete this service request?","It will be removed from the asset's service history. Parts issued from inventory go back in stock.","Delete",()=>{ const m = srById(ds.id); srParts(m).forEach(p=>p.status="Cancelled"); srSync(m); S.maintenance = S.maintenance.filter(x=>x.id!==ds.id); save(); closeModal(); rerender(); }); break;
    case "new-cust": simpleForm("cust"); break;
    case "loc-assets": ui.items.q = ""; ui.items.loc = ds.id; ui.items.locExact = ds.exact ? ds.id : ""; ui.items.tab = "all"; ui.items.cat = ""; go("items"); break;
    case "ctab": ui.cust.tab = ds.v; rerender(); break;
    case "cust-archive": archiveCustomer(ds.id, true); break;
    case "cust-restore": archiveCustomer(ds.id, false); break;
    case "cust-del": deleteCustomer(ds.id); break;
    case "edit-cust": simpleForm("cust", ds.id); break;
    case "new-loc": simpleForm("loc", null, ds.parent); break;
    case "loc-open": ui.locs.open = ds.id; if (route()==="locations") { rerender(); window.scrollTo(0,0); } else go("locations"); break;
    case "loc-close": ui.locs.open = ds.id===undefined ? (locParent(locById(ui.locs.open))?.id||null) : (ds.id||null); rerender(); break;
    case "edit-loc": simpleForm("loc", ds.id); break;
    case "save-simple": saveSimple(); break;
    case "cust-orders": ui.orders.q = custById(ds.id).name; ui.orders.tab="all"; ui.orders.filter=""; go("orders"); break;
    case "loc-orders": ui.orders.q = locById(ds.id).code || locById(ds.id).name; ui.orders.tab="all"; ui.orders.filter=""; go("orders"); break;
    case "add-rate": S.settings[ds.v].push(ds.v==="delivery"?{name:"New zone",rate:0}:{name:"New charge",rate:0,unit:"each"}); save(); rerender(); break;
    case "del-rate": S.settings[ds.v].splice(+ds.i,1); save(); rerender(); break;
    case "file-save": fileSave(); break;
    case "file-open": fileOpen(); break;
    case "reset": confirmBox("Reset all data?","This replaces every order, asset, customer and service record with the original rate sheet and sample orders. Save a data file first if you want to keep anything.","Reset data",()=>{ S = migrate(seed()); fileHandle=null; save(); render(); toast("Data reset"); }); break;
    case "gopen": closeModal(); $("#gresults").hidden = true; $("#gsearch").value=""; if (ds.kind==="order") viewOrder(+ds.v); else if (ds.kind==="item") openAsset(ds.v); else if (ds.kind==="stock") go("asset-stock"); else if (ds.kind==="inv") invAdjust(ds.v); else if (ds.kind==="co") openCustodyCheckin([ds.v]); else if (ds.kind==="cust"){ ui.orders.q=custById(ds.v).name; ui.orders.tab="all"; go("orders"); } else { ui.orders.q=locById(ds.v).code; ui.orders.tab="all"; go("orders"); } break;
  }
});
document.addEventListener("change", e => {
  const t = e.target, ds = t.dataset;
  if (ds.act==="pick"){ if (t.checked) draft.lines.push(lineFrom(ds.id)); else draft.lines = draft.lines.filter(l=>l.itemId!==ds.id); drawOrderForm(); return; }
  if (ds.act==="ci-maint"){ ci.maint[ds.id] = t.checked; return; }
  if (ds.act==="co-pick"){ if (t.checked) codraft.assets.push(ds.id); else codraft.assets = codraft.assets.filter(x=>x!==ds.id); drawCheckoutForm(); return; }
  if (ds.act==="co-sel"){ t.checked?ui.co.sel.add(ds.id):ui.co.sel.delete(ds.id); rerender(); return; }
  if (ds.act==="co-selall"){ [...document.querySelectorAll('[data-act="co-sel"]')].forEach(x=>t.checked?ui.co.sel.add(x.dataset.id):ui.co.sel.delete(x.dataset.id)); rerender(); return; }
  if (ds.act==="coin-maint"){ coin.maint[ds.id] = t.checked; return; }
  if (ds.co==="type"){ codraft.type = t.value; codraft.pq = ""; drawCheckoutForm(); refocus("co-pq"); return; }
  if (ds.co && !["holder","notes","pq"].includes(ds.co)){ codraft[ds.co] = t.value; drawCheckoutForm(); refocus(t.id); return; }
  if (ds.ciLoc){ ci.loc[ds.ciLoc] = t.value; return; }
  if (ds.coinLoc){ coin.loc[ds.coinLoc] = t.value; return; }
  if (ds.act==="cmb-pick"){ cmb.picked = t.checked ? [...new Set(cmb.picked.concat(t.value))] : cmb.picked.filter(x=>x!==t.value); drawCombine(); return; }
  if (ds.act==="cmb-mode"){ cmb.mode = t.value; drawCombine(); return; }
  if (ds.cmb==="parent"||ds.cmb==="into"){ cmb[ds.cmb] = t.value; drawCombine(); return; }
  if (ds.coinRest){ coin.rest[ds.coinRest] = t.value; return; }
  if (ds.coin){ coin[ds.coin] = t.value; return; }
  if (ds.mp!=null){ const p = mdraft.parts[+ds.mp]; if (ds.f==="invId"){ p.invId=t.value; p.cost=+S.inventory.find(x=>x.id===t.value)?.unitCost||0; } else p.qty=+t.value; $("#m-parts").innerHTML = partsEditor(); refocus(t.id); return; }
  if (ds.act==="sel"){ const n=+ds.no; t.checked?ui.orders.sel.add(n):ui.orders.sel.delete(n); rerender(); return; }
  if (ds.act==="selall"){ const rows=[...document.querySelectorAll('[data-act="sel"]')].map(x=>+x.dataset.no); rows.forEach(n=>t.checked?ui.orders.sel.add(n):ui.orders.sel.delete(n)); rerender(); return; }
  if (ds.actChange==="dzone"){ const z = S.settings.delivery.find(d=>d.name===t.value); draft.delivery = z?{name:z.name,rate:z.rate,trips:draft.delivery?.trips||2}:{name:"",rate:0,trips:0}; drawOrderForm(); return; }
  if (ds.of && ["description","notes","newNo"].includes(ds.of)){ setDraft(ds.of, t.value); return; }
  if (ds.of && ds.of!=="pq"){ setDraft(ds.of, t.value); drawOrderForm(); refocus(t.id); return; }
  if (ds.sl!=null && draft){ const x = draft.sales[+ds.sl]; if (ds.f==="invId"){ x.invId = t.value; x.price = +invById(t.value)?.price||0; } else x[ds.f] = +t.value; drawOrderForm(); refocus(t.id); return; }
  if (ds.ch!=null){ const arr = ci?ci.charges:draft.charges; const c = arr[+ds.ch]; if (ds.f==="name"){ c.name=t.value; const s=S.settings.charges.find(x=>x.name===t.value); if (s) c.rate=s.rate; } else c[ds.f]=+t.value; ci?drawCheckin():drawOrderForm(); refocus(t.id); return; }
  if (ds.ci){ ci[ds.ci] = ds.ci==="billActual" ? t.value==="1" : t.value; drawCheckin(); refocus(t.id); return; }
  if (ds.nextorder){ const v = +t.value; if (me()?.role!=="Admin" || !Number.isInteger(v) || v < 1){ toast("Next order # has to be a whole number, 1 or higher"); rerender(); return; } S.nextOrder = v; nextFreeOrderNo(); save(); rerender(); toast(`Next order will be #${S.nextOrder}`); return; }
  if ((ds.set || ds.rate) && !can("setup")) { toast("Your role can't change settings"); rerender(); return; }
  if (t.id==="us-role"){ const h=$("#us-role-desc"); if (h) h.textContent = ROLES[t.value].desc; const w = $("#us-cust-wrap"); if (w) w.hidden = t.value!=="Superintendent"; return; }
  if (ds.set){ S.settings[ds.set] = (ds.set==="cycle"||ds.set==="taxRate"||ds.set==="shopTargetDays") ? +t.value : t.value; save(); if (ds.set==="contact") renderRail(); toast("Saved"); return; }
  if (ds.rate){ const r = S.settings[ds.rate][+ds.i]; r[ds.f] = ds.f==="rate" ? +t.value : t.value; save(); toast("Saved"); return; }
  if (ds.change==="ostatus"){ ui.orders.status = t.value; rerender(); }
  if (ds.change==="myloc"){ ui.myrep.loc = t.value; rerender(); }
  if (ds.change==="myper"){ ui.myrep.period = t.value; rerender(); }
  if (ds.change==="rqcat"){ ui.req.cat = t.value; rerender(); }
  if (ds.change==="avcat"){ ui.av.cat = t.value; rerender(); }
  if (ds.change==="avall"){ ui.av.all = t.checked; rerender(); }
  if (t.dataset.rqi!=null){ const c = ui.req.cart[+t.dataset.rqi]; if (c){ const v = +t.value; if (v>0) c.qty = v; else t.value = c.qty; saveCart(); reqEstimate(); } }
  if (["rq-need","rq-need-t","rq-until","rq-until-t"].includes(t.id)) reqEstimate();
  if (ds.change==="iuse"){ ui.items.use = t.value; rerender(); }
  if (ds.change==="invsort"){ ui.inv.sort = t.value; rerender(); }
  if (ds.change==="invcat"){ ui.inv.cat = t.value; rerender(); }
  if (ds.change==="msev"){ ui.maint.sev = t.value; rerender(); }
  if (ds.change==="mtype"){ ui.maint.type = t.value; rerender(); }
  if (ds.change && ds.change.startsWith("sr-") && srFieldChange(t)) return;
  if (t.id==="m-op"){ const h = $("#m-op-hint"); if (h) h.textContent = (OP_STATUS.find(x=>x[0]===t.value)||["",""])[1]; }
  if (t.id==="m-sev" && t.value==="High" && $("#m-op") && $("#m-op").value!=="Deadlined"){ $("#m-op").value = "Deadlined"; $("#m-op").dispatchEvent(new Event("change",{bubbles:true})); }
  if (t.id==="m-sev"){ const h = $("#m-sev-hint"), d = SEVERITY.find(x=>x[0]===t.value); if (h) h.textContent = d ? d[1].split(" – ")[1] : ""; }
  if (ds.change==="icat"){ ui.items.cat = t.value; rerender(); }
  if (ds.change==="iloc"){ ui.items.loc = t.value; rerender(); }
  if (ds.change==="rperiod"){ ui.reports.period = t.value; rerender(); }
  if (ds.change==="rloc"){ ui.reports.loc = t.value; rerender(); }
  if (ds.change==="rco"){ ui.reports.withCo = t.checked; rerender(); }
  if (ds.change==="rlate"){ ui.reports.lateOnly = t.checked; rerender(); }
});
let tmr;
document.addEventListener("input", e => {
  const t = e.target, ds = t.dataset;
  if (ds.input==="oq"){ ui.orders.q = t.value; clearTimeout(tmr); tmr = setTimeout(rerender,150); }
  if (ds.input==="avq"){ ui.av.q = t.value; clearTimeout(tmr); tmr = setTimeout(()=>{ rerender(); refocus("av-q"); },200); }
  if (ds.input==="rqq"){ ui.req.q = t.value; clearTimeout(tmr); tmr = setTimeout(()=>{ rerender(); refocus("rq-q"); },200); }
  if (ds.input==="invq"){ ui.inv.q = t.value; clearTimeout(tmr); tmr = setTimeout(()=>{ rerender(); refocus("inv-q"); },200); }
  if (ds.input==="mq"){ ui.maint.q = t.value; clearTimeout(tmr); tmr = setTimeout(()=>{ rerender(); refocus("m-q"); },200); }
  if (ds.input==="iq"){ ui.items.q = t.value; clearTimeout(tmr); tmr = setTimeout(rerender,150); }
  if (ds.of==="pq"){ draft.pq = t.value; clearTimeout(tmr); tmr = setTimeout(()=>{ drawOrderForm(); refocus("f-pq"); },150); }
  if (t.id==="gsearch") globalSearch(t.value);
  if (ds.cmb==="newName") cmb.newName = t.value;
  if (t.id==="m-q"){ const h = searchResultsHtml(t.value); $("#m-results").innerHTML = h || '<div class="small muted">Start typing to search everything.</div>'; }
  if (t.id==="qci-q"){ const q = t.value.trim().toLowerCase(); $("#qci-list").innerHTML = (window._qci||[]).filter(r=>!q||r.q.includes(q)).map(r=>r.html).join("") || '<div class="empty">Nothing matches.</div>'; }
  if (t.id==="it-id" && !$("#itemform")?.dataset.orig){ const s = seriesGroup(t.value), sel = $("#it-cat"); if (s && sel){ if (![...sel.options].some(o=>o.value===s)) sel.add(new Option(s, s)); sel.value = s; } }
  if (ds.of && ["description","notes","newNo"].includes(ds.of)) setDraft(ds.of, t.value);
  if (ds.input==="coq"){ ui.co.q = t.value; clearTimeout(tmr); tmr = setTimeout(rerender,150); }
  if (ds.co==="holder"||ds.co==="notes"){ codraft[ds.co] = t.value; const b=document.querySelector('[data-act="co-save"]'); if (b) b.disabled = !coValid(codraft); }
  if (ds.cq){ const l = (ds.cq==="inv"?codraft.inv:codraft.stock)[+ds.i]; l.qty = +t.value; $("#co-cart").innerHTML = coCart(codraft); refocus(t.id); const b=document.querySelector('[data-act="co-save"]'); if (b) b.disabled = !coValid(codraft); }
  if (ds.coinQ){ coin.back[ds.coinQ] = t.value; const c = S.custody.find(x=>x.id===ds.coinQ), w = $("#cr-wrap-"+c.id); if (w) w.innerHTML = coinRestHtml(c); }
  if (ds.co==="pq"){ codraft.pq = t.value; clearTimeout(tmr); tmr = setTimeout(()=>{ drawCheckoutForm(); refocus("co-pq"); },150); }
  if (ds.coin==="notes") coin.notes = t.value;
});
document.addEventListener("keydown", e => { if (e.key==="Escape") $("#pmenu")?.remove(); if (e.key==="Escape" && $("#modal-root").innerHTML){ if (dwDrag) return;
  if (modalHasForm() && Date.now() - escArmed > 2500){ escArmed = Date.now(); nudgeModal("Press Esc again to close without saving."); return; }
  escArmed = 0; closeModal(); draft=null; ci=null; codraft=null; coin=null; mdraft=null; } });
function refocus(id){ const el = document.getElementById(id); if (el){ el.focus(); try{ const v=el.value; if (el.setSelectionRange && el.type!=="number" && el.type!=="datetime-local") el.setSelectionRange(v.length,v.length);}catch(_){} } }
function setDraft(path, v){ if (path==="delivery.trips") draft.delivery.trips = +v; else if (path==="adjustment") draft.adjustment = +v; else draft[path] = v; }
function searchResultsHtml(v){
  const q = v.trim().toLowerCase(); if (!q) return "";
  const res = [];
  S.orders.filter(o=>orderSearchText(o).includes(q)).slice(0,5).forEach(o=>res.push(["order",o.no,`Order #${o.no}`,`${locLabel(locById(o.locationId))} · ${orderStatus(o)}`]));
  S.items.filter(i=>(i.id+" "+i.name).toLowerCase().includes(q)).slice(0,5).forEach(i=>res.push(["item",i.id,`#${i.id} ${i.name}`,itemStatus(i)]));
  S.customers.filter(c=>!c.deleted && c.name.toLowerCase().includes(q)).slice(0,3).forEach(c=>res.push(["cust",c.id,c.name,c.archived?"Customer · archived":"Customer"]));
  S.locations.filter(l=>locLabel(l).toLowerCase().includes(q)).slice(0,3).forEach(l=>res.push(["loc",l.id,locLabel(l),"Location"]));
  S.stock.filter(s=>s.name.toLowerCase().includes(q)).slice(0,3).forEach(s=>res.push(["stock",s.id,s.name,`${s.qty-stockOut(s.id)} of ${s.qty} available`]));
  S.inventory.filter(i=>!i.retired && (i.name+" "+(i.sku||"")).toLowerCase().includes(q)).slice(0,3).forEach(i=>res.push(["inv",i.id,i.name,`${i.qty} ${i.unit} on hand`]));
  S.custody.filter(c=>!c.inAt && (c.holder+" "+c.ref).toLowerCase().includes(q)).slice(0,3).forEach(c=>res.push(["co",c.id,`${c.ref} · ${custodyLabel(c)}`,`Checked out to ${c.holder}`]));
  return res.length ? res.map(([k,v,t,s])=>`<button data-act="gopen" data-kind="${k}" data-v="${esc(v)}"><span class="k">${{order:"Order",item:"Asset",cust:"Person",loc:"Job",stock:"Stock",inv:"Inventory",co:"Check-out"}[k]}</span><span>${esc(t)}<br><span class="small muted">${esc(s)}</span></span></button>`).join("") : `<div class="empty" style="padding:16px">No matches for “${esc(v)}”</div>`;
}
function globalSearch(v){
  const box = $("#gresults"); if (!v.trim()){ box.hidden = true; return; }
  box.innerHTML = searchResultsHtml(v); box.hidden = false;
}

const VIEWS = {"my-reports":vMyReports,request:vRequestCatalog,"my-requests":vMyRequests,requests:vRequestsAdmin,users:vUsers,dashboard:vDashboard,orders:vOrders,checkout:vCheckout,items:vItems,"asset-stock":vStock,inventory:vInventory,groups:vGroups,maintenance:vMaint,locations:vLocations,customers:vCustomers,reports:vReports,setup:vSetup};
try { const th = localStorage.getItem(KEY+"-theme"); if (th) document.documentElement.dataset.theme = th; } catch(e) {}
/* start from whichever browser copy is newest: IndexedDB holds the full data even when localStorage is too small */
/* boot: see cloud.js */

