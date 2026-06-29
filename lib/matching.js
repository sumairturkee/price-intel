export function normModel(s) {
return s.toLowerCase()
.replace(/\bmotorola\b/g,'moto').replace(/\bsamsung galaxy\b/g,'samsung')
.replace(/[^a-z0-9]/g,' ').replace(/\b(galaxy|series|5g|4g|new|old|mrp)\b/g,' ')
.replace(/\s+/g,' ').trim()
.replace(/(\b[a-z]\d+)\s+([a-z]{1,2}\b)/g,'$1$2')
.replace(/(\b\d+)\s+([a-z]{1,2}\b)/g,'$1$2');
}
export function normModelBase(s) {
return normModel(s).replace(/\b(activated|fresh)\b/g,' ').replace(/\s+/g,' ').trim();
}
function wordSet(s){return new Set(s.split(' ').filter(w=>w.length>0));}
function jaccard(a,b){const wa=wordSet(a),wb=wordSet(b);const i=[...wa].filter(w=>wb.has(w)).length;const u=new Set([...wa,...wb]).size;return u?i/u:0;}
function contains(a,b){const wa=wordSet(a),wb=wordSet(b);if(!wa.size||!wb.size)return 0;const sm=wa.size<=wb.size?wa:wb,lg=wa.size<=wb.size?wb:wa;return[...sm].filter(w=>lg.has(w)).length/sm.size;}
function score(a,b){return Math.max(jaccard(a,b),contains(a,b)*0.9);}
function numCore(w){return w.replace(/[a-z]/g,'');}
function modelCodes(n){return n.split(' ').filter(w=>w.length>1&&/\d/.test(w)).map(numCore).filter(Boolean);}
function smartScore(a,b){const ca=modelCodes(a),cb=modelCodes(b);if(ca.length>0&&cb.length>0&&!ca.some(c=>cb.includes(c)))return 0;return score(a,b);}
export function hasStockTag(s){return /\b(activated|fresh)\b/i.test(s);}
function normVariant(v){return(v||'').toLowerCase().replace(/[^a-z0-9]/g,'');}
function variantMatches(a,b){const ca=normVariant(a),cb=normVariant(b);if(!ca||!cb)return true;return ca===cb;}
function getBrand(m){return m.toLowerCase().replace(/motorola/,'moto').split(/[\s(]/)[0];}
const BRANDS=['samsung','realme','poco','redmi','xiaomi','moto','motorola','vivo','oppo','iqoo','narzo','lava','alcatel','iphone','apple','nokia'];
function brandMismatch(a,b){const ba=getBrand(a),bb=getBrand(b);if(BRANDS.includes(ba)&&BRANDS.includes(bb)&&ba!==bb)return true;return false;}
export function findMyMatch(compModel,compVariant,myList){
const compIsActiv=/\bactivated\b/i.test(compModel);
const normComp=normModel(compModel);
let best=null,bestScore=0;
for(const m of myList){
if(brandMismatch(compModel,m.model))continue;
const myIsActiv=/\bactivated\b/i.test(m.model);
const normMy=normModel(m.model);
let s=0;
if(compIsActiv){if(myIsActiv)s=smartScore(normComp,normMy);else if(!hasStockTag(m.model))s=smartScore(normComp,normMy)*0.3;else s=0;}
else{if(myIsActiv)s=0;else s=smartScore(normComp,normMy);}
if(s>=0.4&&!variantMatches(compVariant,m.variant))s=0;
if(s>bestScore){bestScore=s;best=m;}
}
return bestScore>=0.4?best:null;
}
