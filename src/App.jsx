import { useState, useEffect, useMemo, useCallback } from "react";
import {
  Lock, Plus, Minus, Trash2, Pencil, Download, Search, ClipboardList,
  Settings, FileText, ArrowLeft, AlertCircle, CheckCircle2, X,
  RefreshCw, Users, Package, Store, Building2, User, ChevronDown, RotateCcw, PlayCircle, StopCircle
} from "lucide-react";
import * as XLSX from "xlsx";
import { safeGet, safeSet, safeList } from "./api.js";

/* ==================================================================== */
/* DATOS BASE                                                            */
/* ==================================================================== */

const CARTIMEX_AREAS_DEFAULT = [
  { id: "COMPRAS", label: "Compras" },
  { id: "CONTROLLING", label: "Controlling" },
  { id: "CYC", label: "CyC" },
  { id: "TESORERIA", label: "Tesorería" },
  { id: "RRHH", label: "RR.HH." },
  { id: "CONTABILIDAD", label: "Contabilidad" },
  { id: "GERENCIA", label: "Gerencia" },
  { id: "CORPORATIVO", label: "Corporativo" },
  { id: "GUARDIANIA", label: "Guardianía" },
  { id: "VENTAS", label: "Ventas" },
  { id: "LOGISTICA", label: "Logística" },
  { id: "ENSAMBLE", label: "Ensamble" },
  { id: "SERVITECH", label: "Servitech" },
  { id: "MARKETING", label: "Marketing" },
];

const CARTIMEX_ITEMS_RAW = [
  [1,"UTIMPOR","ADHESIVAS MULTIPEG (FUNDITA)","UND"],[2,"UTIMPOR","ALMOHADILLA PARA DEDO","UND"],
  [3,"UTIMPOR","ARCHIVADOR UNICO NEGRO T/MEMO","UND"],[4,"UTIMPOR","ARCHIVADOR UNICO NEGRO T/OFICIO","UND"],
  [5,"UTIMPOR","ARCHIVADOR UNICO NEGRO T/OFICIO LOMO FINO","UND"],[6,"CARTIMEX","BLOCK DE VALE DE CAJA","UND"],
  [7,"UTIMPOR","BORRADOR DE PIZARRA ACRILICO","UND"],[8,"UTIMPOR","BORRADOR PELIKAN PZ 20","UND"],
  [9,"UTIMPOR","CAJA VINCHAS ALEX","UND"],[10,"UTIMPOR","CALCULADORAS CASIO","UND"],
  [11,"UTIMPOR","CERA CONTAR SORTKWK","UND"],[12,"UTIMPOR","CINTA MASKINGTAPE 31/4 x 25 YDAS","UND"],
  [13,"UTIMPOR","CINTA SCOOT 1/2X24 YDS TUBO","UND"],[14,"UTIMPOR","CINTAS DE EMPAQUE TRANSPARENTE","UND"],
  [15,"UTIMPOR","CLIPS ALEX ESTÁNDAR","UND"],[16,"UTIMPOR","CLIPS ALEX MARIPOSA","UND"],
  [17,"UTIMPOR","CUADERNOS ACADEMICOS CUADRICULADO 100 HOJAS","UND"],[18,"CARTIMEX","CUADERNO CONTROL DE INGRESO","UND"],
  [19,"UTIMPOR","DISPENSADOR MEDIANO CINTA","UND"],[20,"UTIMPOR","ESFEROGRAFICAS COLOR AZUL","UND"],
  [21,"UTIMPOR","ESFEROGRAFICAS COLOR NEGRO","UND"],[22,"UTIMPOR","ESFEROGRAFICAS COLOR ROJO","UND"],
  [23,"UTIMPOR","ESTILETE GRANDE","UND"],[24,"UTIMPOR","FOLDER MANILA IDEAL","UND"],
  [25,"CARTIMEX","FUNDA PORTAPAPEL F4","UND"],[26,"UTIMPOR","GOMA EN BARRA KW 36 GRMS","UND"],
  [27,"UTIMPOR","GOMA LIQUIDA BIOPLAST","UND"],[28,"UTIMPOR","GRAPA 26/6","UND"],
  [29,"UTIMPOR","GRAPADORA METAL MEDIANA","UND"],[30,"UTIMPOR","LAPIZ ARTESCO 2HB","UND"],
  [31,"UTIMPOR","LIBRETAS TAQUIGRAFICAS CUADROS","UND"],[32,"UTIMPOR","LIGAS FUNDAS","UND"],
  [33,"UTIMPOR","LIQUIDPAPER","UND"],[34,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO AZUL","UND"],
  [35,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO NEGRO","UND"],[36,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO ROJO","UND"],
  [37,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO VERDE","UND"],[38,"UTIMPOR","MARCADOR PUNTA FINA COLOR AZUL","UND"],
  [39,"UTIMPOR","MARCADOR PUNTA FINA COLOR NEGRO","UND"],[40,"UTIMPOR","MARCADOR PUNTA FINA COLOR ROJO","UND"],
  [41,"UTIMPOR","MARCADOR PERMANENTE AZUL","UND"],[42,"UTIMPOR","MARCADOR PERMANENTE NEGRO","UND"],
  [43,"UTIMPOR","MARCADOR PERMANENTE ROJO","UND"],[44,"UTIMPOR","MARCADORES DETECTORES DE BILLETES FALSOS","UND"],
  [45,"CARTIMEX","PAPELERA METALICA","UND"],[46,"UTIMPOR","PERFORADORA MEDIANA","UND"],
  [47,"UTIMPOR","PORTA CLIP","UND"],[48,"UTIMPOR","PORTALAPIZ","UND"],
  [49,"UTIMPOR","POST IT 76 MM X 76 MM MEDIANO","UND"],[50,"UTIMPOR","POST IT PEQUEÑO","UND"],
  [51,"UTIMPOR","REFUERZOS PARA HOJAS (HOJALILLOS)","UND"],[52,"UTIMPOR","REGLA 30 CM","UND"],
  [53,"UTIMPOR","REPUESTO PARA ESTILETE","UND"],[54,"UTIMPOR","RESALTADORES AMARILLOS","UND"],
  [55,"UTIMPOR","RESALTADORES NARANJA","UND"],[56,"UTIMPOR","RESALTADORES ROSADO","UND"],
  [57,"UTIMPOR","RESALTADORES VERDE","UND"],[58,"UTIMPOR","SACAGRAPA","UND"],
  [59,"UTIMPOR","SACAPUNTA METALICO PEQUEÑO","UND"],[60,"CARTIMEX","SEPARADORES PLASTICOS TAMAÑO A4","UND"],
  [61,"UTIMPOR","SOBRE BOND IDEAL T/CARTA 60GRS","UND"],[62,"UTIMPOR","SOBRE BOND IDEAL T/OFICIO 60GRS","UND"],
  [63,"UTIMPOR","SOBRE MANILA TAMAÑO EXTRAGRANDE","UND"],[64,"UTIMPOR","SOBRE MANILA TAMAÑO OFICIO","UND"],
  [65,"UTIMPOR","SOBRE MANILA TAMAÑO PEQUEÑO","UND"],[66,"UTIMPOR","TABLERO ACERO ARTESCO OFICIO","UND"],
  [67,"UTIMPOR","TIJERA METALICA PARA OFICINA","UND"],[68,"UTIMPOR","TINTA PARA SELLOS PELIKAN (AZUL)","UND"],
  [69,"UTIMPOR","TINTA PARA SELLOS PELIKAN (NEGRA)","UND"],[70,"CARTIMEX","CARPETAS MEMBRETADAS CARTIMEX","UND"],
  [71,"CARTIMEX","HOJAS MEMBRETADAS DE CARTIMEX","UND"],[72,"CARTIMEX","HOJAS MEMBRETADAS DE COMPUTRONSA","UND"],
  [73,"CARTIMEX","RESMAS DE PAPEL","UND"],[74,"CARTIMEX","SOBRES MEMBRETADOS COMPUTRON","UND"],
  [75,"CARTIMEX","SOBRES MEMBRETADOS CARTIMEX","UND"],[76,"CALBAQ","PILAS AA","UND"],
  [77,"CALBAQ","PILAS AAA","UND"],
];
const CARTIMEX_ITEMS_DEFAULT = CARTIMEX_ITEMS_RAW.map(([numero, proveedor, item, unidad]) => ({
  id: `itm${numero}`, numero, proveedor, item, unidad,
}));

const COMPUTRON_TIENDAS_LABELS = [
  "PAP CALIFORNIA","PAP DURAN","PAP FLORIDA","PAP PORTETE","KENNEDY","CEIBOS",
  "CEIBOS CREDITO Y COBRANZA","CENTRO 2","RIOCENTRO SUR","RIOCENTRO EL DORADO","PASEO DAULE",
  "PORTETE","CALIFORNIA","DURAN","QUEVEDO","MACHALA","PIAZZA MACHALA","LA LIBERTAD","MILAGRO",
  "MALL DEL SUR","CENTRO3","BABAHOYO","DAULE","MALL DEL NORTE","QUITO NORTE","QUICENTRO SUR",
  "CUMBAYA","EL RECREO","SANTO DOMINGO","RIOBAMBA SHOPPING","QUICENTRO SUR 2","SANGOLQUI",
  "CONDADO SHOPPING","LATACUNGA","OTAVALO","AMBATO CENTRO","ESMERALDAS","SANGOLQUI 2",
];

const COMPUTRON_ITEMS_OFICINA_RAW = [
  [1,"UTIMPOR","ADHESIVAS MULTIPEG (FUNDITA)","PAQUETES"],[2,"UTIMPOR","ARCHIVADOR ECONOMICO NEGRO T/MEMO","UNIDADES"],
  [3,"UTIMPOR","ARCHIVADOR UNICO NEGRO T/OFICIO","UNIDADES"],[4,"UTIMPOR","BORRADOR DE PIZARRA ACRILICO","UNIDADES"],
  [5,"UTIMPOR","BORRADOR PELIKAN PZ 20","UNIDADES"],[6,"UTIMPOR","CAJA VINCHAS ALEX 50 UNIDADES","CAJAS"],
  [7,"UTIMPOR","CALCULADORAS CASIO","UNIDADES"],[8,"UTIMPOR","CERA CONTAR SORTKWK","UNIDADES"],
  [9,"UTIMPOR","CINTA MASKINGTAPE 31/4 x 25 YDAS","UNIDADES"],[10,"UTIMPOR","CINTA SCOOT 1/2X24 YDS TUBO","UNIDADES"],
  [11,"UTIMPOR","CINTAS DE EMPAQUE TRANSPARENTE","UNIDADES"],[12,"UTIMPOR","CLIPS ALEX ESTÁNDAR x 100","CAJAS"],
  [13,"UTIMPOR","CLIPS ALEX MARIPOSA X 50","CAJAS"],[14,"UTIMPOR","CUADERNO CUADRICULADO JUNIOR 100 HOJAS","UNIDADES"],
  [15,"UTIMPOR","CUADERNOS ACADEMICOS CUADRICULADO 100 HOJAS","UNIDADES"],[16,"UTIMPOR","DISPENSADOR MEDIANO CINTA","UNIDADES"],
  [17,"UTIMPOR","ESFEROGRAFICAS COLOR AZUL","UNIDADES"],[18,"UTIMPOR","ESFEROGRAFICAS COLOR NEGRO","UNIDADES"],
  [19,"UTIMPOR","ESFEROGRAFICAS COLOR ROJO","UNIDADES"],[20,"UTIMPOR","ESTILETE GRANDE KENDO","UNIDADES"],
  [21,"UTIMPOR","FOLDER MANILA IDEAL","UNIDADES"],[22,"CARTIMEX","FUNDA PORTAPAPEL F4 X10","PAQUETES"],
  [23,"UTIMPOR","GOMA EN BARRA KW 36 GRMS","UNIDADES"],[24,"UTIMPOR","GOMA LIQUIDA BIOPLAST","UNIDADES"],
  [25,"UTIMPOR","GRAPA 26/6 ARTESCO","UNIDADES"],[26,"UTIMPOR","GRAPADORA METAL MEDIANA","UNIDADES"],
  [27,"CARTIMEX","HOJAS MEMBRETADAS DE COMPUTRONSA","UNIDADES"],[28,"CARTIMEX","SOBRES MEMBRETADAS DE COMPUTRONSA","UNIDADES"],
  [29,"UTIMPOR","LAPIZ ARTESCO 2HB","UNIDADES"],[30,"UTIMPOR","LIGAS FUNDAS","PAQUETES"],
  [31,"UTIMPOR","LIQUIDPAPER","UNIDADES"],[32,"UTIMPOR","MARCADOR PUNTA FINA COLOR NEGRO","UNIDADES"],
  [33,"UTIMPOR","MARCADOR PUNTA FINA COLOR ROJO","UNIDADES"],[34,"UTIMPOR","MARCADOR PUNTA FINA COLOR AZUL","UNIDADES"],
  [35,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO AZUL","UNIDADES"],[36,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO NEGRO","UNIDADES"],
  [37,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO ROJO","UNIDADES"],[38,"UTIMPOR","MARCADOR DE PIZARRA ACRILICO VERDE","UNIDADES"],
  [39,"UTIMPOR","MARCADOR PERMANENTE ROJO","UNIDADES"],[40,"UTIMPOR","MARCADOR PERMANENTE AZUL","UNIDADES"],
  [41,"UTIMPOR","MARCADOR PERMANENTE NEGRO","UNIDADES"],[42,"UTIMPOR","MARCADORES DETECTORES DE BILLETES FALSOS","UNIDADES"],
  [43,"CARTIMEX","PAPELERA METALICA 2 SERVICIOS","UNIDADES"],[44,"UTIMPOR","PERFORADORA MEDIANA","UNIDADES"],
  [45,"UTIMPOR","PORTA CLIP","UNIDADES"],[46,"UTIMPOR","PORTALAPIZ","UNIDADES"],
  [47,"UTIMPOR","POST IT 76 MM X 76 MM MEDIANO","UNIDADES"],[48,"UTIMPOR","POST IT PEQUEÑO","UNIDADES"],
  [49,"CALBAQ","PILAS AA","PAQUETES"],[50,"CALBAQ","PILAS AAA","PAQUETES"],
  [51,"UTIMPOR","REFUERZOS PARA HOJAS (HOJALILLOS)","PAQUETES"],[52,"UTIMPOR","REGLA 30 CM","UNIDADES"],
  [53,"UTIMPOR","REPUESTO PARA ESTILETE X10","PAQUETES"],[54,"UTIMPOR","RESALTADORES AMARILLOS","UNIDADES"],
  [55,"UTIMPOR","RESALTADORES NARANJA","UNIDADES"],[56,"UTIMPOR","RESALTADORES ROSADO","UNIDADES"],
  [57,"UTIMPOR","RESALTADORES VERDE","UNIDADES"],[58,"UTIMPOR","SACAGRAPA","UNIDADES"],
  [59,"UTIMPOR","SACAPUNTA METALICO PEQUEÑO","UNIDADES"],[60,"CARTIMEX","SEPARADORES PLASTICOS TAMAÑO A4","PAQUETES"],
  [61,"UTIMPOR","SOBRE BOND IDEAL T/CARTA 60GRS","UNIDADES"],[62,"UTIMPOR","SOBRE BOND IDEAL T/OFICIO 60GRS","UNIDADES"],
  [63,"UTIMPOR","SOBRE MANILA TAMAÑO EXTRAGRANDE A3","UNIDADES"],[64,"UTIMPOR","SOBRE MANILA TAMAÑO OFICIO A4","UNIDADES"],
  [65,"UTIMPOR","SOBRE MANILA TAMAÑO PEQUEÑO A6","UNIDADES"],[66,"UTIMPOR","TABLERO ACERO ARTESCO OFICIO","UNIDADES"],
  [67,"UTIMPOR","TIJERA METALICA PARA OFICINA","UNIDADES"],[68,"UTIMPOR","TINTA PARA SELLOS PELIKAN (AZUL)","UNIDADES"],
  [69,"UTIMPOR","TINTA PARA SELLOS PELIKAN (NEGRA)","UNIDADES"],[70,"UTIMPOR","TINTA PARA SELLOS PELIKAN (ROJA)","UNIDADES"],
];
const COMPUTRON_ITEMS_OFICINA_DEFAULT = COMPUTRON_ITEMS_OFICINA_RAW.map(([numero, proveedor, item, unidad]) => ({
  id: `ofc${numero}`, numero, proveedor, item, unidad,
}));

const COMPUTRON_ITEMS_LIMPIEZA_RAW = [
  [1,"DAAS","ALGODÓN NATURAL PAQUETE 70 GRS","PAQUETES"],[2,"CALBAQ","CEPILLO MANO DE OSO PARA SANITARIOS","UNIDADES"],
  [3,"MERGAMA","DETERGENTE 1 KG.","UNIDADES"],[4,"CALBAQ","ESCOBA JOLLY PLUMADA MANGO PLASTIFICADO","UNIDADES"],
  [5,"MERGAMA","ESPONJA SALVAUÑAS SAPOLIO","UNIDADES"],[6,"MERGAMA","FUNDA DE BASURA 18 X 20 (10 UNIDADES)","PAQUETES"],
  [7,"MERGAMA","FUNDA DE BASURA 30 X 36 (10 UNIDADES)","PAQUETES"],[8,"MERGAMA","FUNDA DE BASURA 36 X 40 (10 UNIDADES)","PAQUETES"],
  [9,"MERGAMA","FUNDA DE BASURA 39 1/2 X 55 (10 UNIDADES)","PAQUETES"],[10,"UNILIMPIO","GALON DE ALCOHOL GEL","UNIDADES"],
  [11,"UNILIMPIO","GALON DE ALCOHOL LIQUIDO","UNIDADES"],[12,"UNILIMPIO","GALON DE JABON LIQUIDO","UNIDADES"],
  [13,"MERGAMA","GAMUZA VILEDA LIMPIAVIDRIOS","UNIDADES"],[14,"MERGAMA","GUANTES BEST CHEM MASTER TALLA 8","UNIDADES"],
  [15,"MERGAMA","GUANTES BEST CHEM MASTER TALLA 9","UNIDADES"],[16,"MERGAMA","GUANTES BEST CHEM MASTER TALLA 7","UNIDADES"],
  [17,"MERGAMA","GUANTES DE EXAMINACION MEDIUM 100 UNIDADES","PAQUETES"],[18,"MERGAMA","INSECTICIDA SPRAY MATA CUCARACHA Y HORMIGAS","UNIDADES"],
  [19,"MERGAMA","INSECTICIDA SPRAY MATA MOSQUITOS Y MOSCAS","UNIDADES"],[20,"MERGAMA","LEJIA SELLO ROJO","UNIDADES"],
  [21,"MERGAMA","LIQUIDO LIMPIA VIDRIO GALON","UNIDADES"],[22,"MERGAMA","PAÑO MICROFIBRA","UNIDADES"],
  [23,"MERGAMA","PAPEL HIGIENICO JUMBO CLASICO X 250 MTRS D.H.","UNIDADES"],[24,"MERGAMA","RECOGEDOR DE BASURA","UNIDADES"],
  [25,"MERGAMA","TACHO MALLADO","UNIDADES"],[26,"CALBAQ","TIPS AMBIENTAL EN PASTILLAS DE 95 GRS.","UNIDADES"],
  [27,"CALBAQ","TIPS BRILLO LAVA VAJILLAS EN CREMA 900G","UNIDADES"],[28,"CALBAQ","TIPS BRILLO LAVA VAJILLAS LIQUIDO 1000ML","UNIDADES"],
  [29,"CALBAQ","TIPS CLORO LIQUIDO LIMPIEZA CITRICA 1000 ML","UNIDADES"],[30,"CALBAQ","TIPS DESINFECTANTE ANTIBACTERIAL LAVANDA LITRO","UNIDADES"],
  [31,"CALBAQ","TIPS DESINFECTANTE ANTIBACTERIAL MANZANA CANELA LITRO","UNIDADES"],[32,"CALBAQ","TIPS PAÑOS DESINFECTANTES LIMON 50 UNIDADES","PAQUETES"],
  [33,"CALBAQ","TIPS TANQUE ECONOPACK DISPLAY X 12 X 48 GR","UNIDADES"],[34,"UNILIMPIO","TOALLA DE MANOS EN ROLLO BLANCA 100MT X4","UNIDADES"],
  [35,"MERGAMA","TRAPEADOR VILEDA","UNIDADES"],
];
const COMPUTRON_ITEMS_LIMPIEZA_DEFAULT = COMPUTRON_ITEMS_LIMPIEZA_RAW.map(([numero, proveedor, item, unidad]) => ({
  id: `lmp${numero}`, numero, proveedor, item, unidad,
}));

const DEFAULT_PASSWORD = "suministros2026";
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const THEME = {
  CARTIMEX: {
    stampBorder: "border-emerald-700 text-emerald-800",
    subtitleText: "text-emerald-800",
    primary: "bg-emerald-800",
    primaryHover: "hover:bg-emerald-700",
    ring: "focus:ring-emerald-600 focus:border-emerald-600",
    activeCardBorder: "border-emerald-400",
    activeCardBg: "bg-emerald-50/70",
    badgeBg: "bg-emerald-100",
    badgeText: "text-emerald-800",
    tabActiveBorder: "border-emerald-700",
    selectedText: "text-emerald-700",
  },
  COMPUTRON: {
    stampBorder: "border-blue-800 text-blue-900",
    subtitleText: "text-blue-900",
    primary: "bg-blue-900",
    primaryHover: "hover:bg-blue-800",
    ring: "focus:ring-blue-700 focus:border-blue-700",
    activeCardBorder: "border-blue-400",
    activeCardBg: "bg-blue-50/70",
    badgeBg: "bg-blue-100",
    badgeText: "text-blue-900",
    tabActiveBorder: "border-blue-800",
    selectedText: "text-blue-800",
  },
};

/* ==================================================================== */
/* STORAGE HELPERS                                                       */
/* ==================================================================== */

const CONFIG_KEY = "config";
const CARTIMEX_AREAS_KEY = "cartimex-areas";
const CARTIMEX_ITEMS_KEY = "cartimex-items";
const COMPUTRON_TIENDAS_KEY = "computron-tiendas";
const COMPUTRON_ITEMS_OFICINA_KEY = "computron-items-oficina";
const COMPUTRON_ITEMS_LIMPIEZA_KEY = "computron-items-limpieza";
const PERIODS_KEY = "periods-list";

const submissionKey = (format, periodSlug, entityId) => `submission:${format}:${periodSlug}:${entityId}`;

const slugify = (s) =>
  (s || "").trim().toUpperCase().replace(/[^\w\-]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");

function makeId(existingIds, label, prefix = "") {
  let base = prefix + (slugify(label) || `ID-${Date.now()}`);
  let id = base, i = 2;
  while (existingIds.includes(id)) { id = `${base}-${i}`; i++; }
  return id;
}

function parseJSON(raw, fallback) {
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch { return fallback; }
}

/* ==================================================================== */
/* UI HELPERS                                                            */
/* ==================================================================== */

function Stamp({ children, theme }) {
  const cls = theme ? theme.stampBorder : "border-amber-600 text-amber-700";
  return (
    <div className={`inline-flex items-center justify-center w-14 h-14 rounded-full border-2 ${cls} font-bold text-[10px] leading-tight text-center rotate-[-6deg] tracking-wide select-none shrink-0`}>
      {children}
    </div>
  );
}
function Perforation() {
  return (
    <div className="flex gap-1.5 px-4 py-2 overflow-hidden">
      {Array.from({ length: 60 }).map((_, i) => <span key={i} className="w-1.5 h-1.5 rounded-full bg-stone-200 shrink-0" />)}
    </div>
  );
}
function LoadingScreen() {
  return (
    <div className="min-h-[500px] flex items-center justify-center bg-stone-50">
      <div className="flex flex-col items-center gap-3 text-stone-500">
        <RefreshCw className="w-6 h-6 animate-spin" />
        <span className="text-sm font-medium">Cargando sistema...</span>
      </div>
    </div>
  );
}

/* ==================================================================== */
/* APP                                                                    */
/* ==================================================================== */

export default function App() {
  const [loading, setLoading] = useState(true);
  const [config, setConfig] = useState(null);
  const [cartimexAreas, setCartimexAreas] = useState([]);
  const [cartimexItems, setCartimexItems] = useState([]);
  const [computronTiendas, setComputronTiendas] = useState([]);
  const [computronItemsOficina, setComputronItemsOficina] = useState([]);
  const [computronItemsLimpieza, setComputronItemsLimpieza] = useState([]);
  const [view, setView] = useState("form");

  useEffect(() => {
    (async () => {
      // Todas las lecturas iniciales en paralelo (antes eran 6-7 peticiones
      // en fila, una esperando a la otra — esto acelera bastante la carga).
      const [cfgRaw, areasRaw, cItemsRaw, legacyItemsRaw, tiendasRaw, ofcRaw, lmpRaw] = await Promise.all([
        safeGet(CONFIG_KEY, true),
        safeGet(CARTIMEX_AREAS_KEY, true),
        safeGet(CARTIMEX_ITEMS_KEY, true),
        safeGet("items-list", true), // migración desde versión antigua
        safeGet(COMPUTRON_TIENDAS_KEY, true),
        safeGet(COMPUTRON_ITEMS_OFICINA_KEY, true),
        safeGet(COMPUTRON_ITEMS_LIMPIEZA_KEY, true),
      ]);

      // config (con migración desde versión anterior de una sola marca)
      let cfg = parseJSON(cfgRaw, null);
      if (!cfg) {
        cfg = { activeFormat: "", periodSlug: "", periodLabel: "", adminPassword: DEFAULT_PASSWORD };
        await safeSet(CONFIG_KEY, JSON.stringify(cfg), true);
      } else if (cfg.activeFormat === undefined) {
        cfg = {
          activeFormat: cfg.currentPeriodSlug ? "CARTIMEX" : "",
          periodSlug: cfg.currentPeriodSlug || "",
          periodLabel: cfg.currentPeriodLabel || "",
          adminPassword: cfg.adminPassword || DEFAULT_PASSWORD,
        };
        await safeSet(CONFIG_KEY, JSON.stringify(cfg), true);
      }
      setConfig(cfg);

      // cartimex areas
      let areas = parseJSON(areasRaw, null);
      if (!areas) { areas = CARTIMEX_AREAS_DEFAULT; await safeSet(CARTIMEX_AREAS_KEY, JSON.stringify(areas), true); }
      setCartimexAreas(areas);

      // cartimex items (con migración desde 'items-list')
      let citems = parseJSON(cItemsRaw, null);
      if (!citems) {
        citems = parseJSON(legacyItemsRaw, null) || CARTIMEX_ITEMS_DEFAULT;
        await safeSet(CARTIMEX_ITEMS_KEY, JSON.stringify(citems), true);
      }
      setCartimexItems(citems);

      // computron tiendas
      let tiendas = parseJSON(tiendasRaw, null);
      if (!tiendas) {
        tiendas = COMPUTRON_TIENDAS_LABELS.map((label) => ({ id: slugify(label), label }));
        await safeSet(COMPUTRON_TIENDAS_KEY, JSON.stringify(tiendas), true);
      }
      setComputronTiendas(tiendas);

      // computron items oficina
      let ofc = parseJSON(ofcRaw, null);
      if (!ofc) { ofc = COMPUTRON_ITEMS_OFICINA_DEFAULT; await safeSet(COMPUTRON_ITEMS_OFICINA_KEY, JSON.stringify(ofc), true); }
      setComputronItemsOficina(ofc);

      // computron items limpieza
      let lmp = parseJSON(lmpRaw, null);
      if (!lmp) { lmp = COMPUTRON_ITEMS_LIMPIEZA_DEFAULT; await safeSet(COMPUTRON_ITEMS_LIMPIEZA_KEY, JSON.stringify(lmp), true); }
      setComputronItemsLimpieza(lmp);

      setLoading(false);
    })();
  }, []);

  const persist = useCallback((setter, key) => async (next) => {
    setter(next);
    await safeSet(key, JSON.stringify(next), true);
  }, []);

  const persistConfig = persist(setConfig, CONFIG_KEY);
  const persistCartimexAreas = persist(setCartimexAreas, CARTIMEX_AREAS_KEY);
  const persistCartimexItems = persist(setCartimexItems, CARTIMEX_ITEMS_KEY);
  const persistComputronTiendas = persist(setComputronTiendas, COMPUTRON_TIENDAS_KEY);
  const persistComputronItemsOficina = persist(setComputronItemsOficina, COMPUTRON_ITEMS_OFICINA_KEY);
  const persistComputronItemsLimpieza = persist(setComputronItemsLimpieza, COMPUTRON_ITEMS_LIMPIEZA_KEY);

  if (loading) return <LoadingScreen />;

  const dataBundle = {
    cartimexAreas, cartimexItems, computronTiendas, computronItemsOficina, computronItemsLimpieza,
  };

  return (
    <div className="min-h-[600px] bg-stone-50 text-stone-800 font-sans">
      {view === "form" && (
        <FormView config={config} data={dataBundle} onGoAdmin={() => setView("adminGate")} />
      )}
      {view === "adminGate" && (
        <AdminGate config={config} onBack={() => setView("form")} onSuccess={() => setView("admin")} />
      )}
      {view === "admin" && (
        <AdminView
          config={config} data={dataBundle}
          onConfigChange={persistConfig}
          onCartimexAreasChange={persistCartimexAreas}
          onCartimexItemsChange={persistCartimexItems}
          onComputronTiendasChange={persistComputronTiendas}
          onComputronItemsOficinaChange={persistComputronItemsOficina}
          onComputronItemsLimpiezaChange={persistComputronItemsLimpieza}
          onExit={() => setView("form")}
        />
      )}
    </div>
  );
}

/* ==================================================================== */
/* FORM VIEW (jefes de tienda / departamento)                            */
/* ==================================================================== */

function PersonalInfoFields({ nombre, apellido, correo, setNombre, setApellido, setCorreo, correoError, theme }) {
  const ring = theme ? theme.ring : "focus:ring-amber-500 focus:border-amber-500";
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
      <div>
        <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide">Nombre</label>
        <input value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Nombre"
          className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 ${ring}`} />
      </div>
      <div>
        <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide">Apellido</label>
        <input value={apellido} onChange={(e) => setApellido(e.target.value)} placeholder="Apellido"
          className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 ${ring}`} />
      </div>
      <div className="sm:col-span-2">
        <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide">Correo electrónico</label>
        <input value={correo} onChange={(e) => setCorreo(e.target.value)} placeholder="correo@empresa.com" type="email"
          className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 ${ring}`} />
        {correoError && <p className="text-xs text-red-600 mt-1">{correoError}</p>}
      </div>
    </div>
  );
}

function ItemCatalog({ items, quantities, setQty, theme }) {
  const t = theme || THEME.CARTIMEX;
  const [search, setSearch] = useState("");

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return items;
    return items.filter((it) => it.item.toLowerCase().includes(q) || String(it.numero).includes(q));
  }, [items, search]);

  const selectedCount = items.filter((it) => (parseInt(quantities[it.id], 10) || 0) > 0).length;

  const bump = (it, delta) => {
    const current = parseInt(quantities[it.id], 10) || 0;
    setQty(it.id, String(Math.max(0, current + delta)));
  };

  return (
    <div>
      <div className="relative">
        <Search className="w-4 h-4 text-stone-400 absolute left-3 top-1/2 -translate-y-1/2" />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar artículo..."
          className={`w-full border border-stone-300 rounded-lg pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 ${t.ring}`} />
      </div>

      {selectedCount > 0 && (
        <p className={`text-xs mt-3 font-medium ${t.selectedText}`}>{selectedCount} artículo{selectedCount !== 1 ? "s" : ""} seleccionado{selectedCount !== 1 ? "s" : ""} en esta lista</p>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 mt-3">
        {filtered.map((it) => {
          const qty = parseInt(quantities[it.id], 10) || 0;
          const active = qty > 0;
          return (
            <div key={it.id} className={`border rounded-xl p-3 flex flex-col gap-2.5 transition-colors ${active ? `${t.activeCardBorder} ${t.activeCardBg}` : "border-stone-200 bg-white"}`}>
              <div className="flex items-start justify-between gap-2">
                <p className="text-sm font-medium text-stone-800 leading-snug">{it.item}</p>
                <span className="text-[10px] font-mono text-stone-300 shrink-0 mt-0.5">#{it.numero}</span>
              </div>
              <span className={`inline-flex w-fit items-center text-[10px] font-bold uppercase tracking-wide rounded-full px-2 py-0.5 ${t.badgeBg} ${t.badgeText}`}>
                {it.unidad}
              </span>
              <div className="flex items-center gap-1.5">
                <button type="button" onClick={() => bump(it, -1)} disabled={!qty}
                  className="w-7 h-7 rounded-md border border-stone-300 flex items-center justify-center text-stone-500 hover:bg-stone-100 disabled:opacity-30 disabled:hover:bg-white shrink-0">
                  <Minus className="w-3.5 h-3.5" />
                </button>
                <input type="number" min="0" inputMode="numeric" value={quantities[it.id] ?? ""}
                  onChange={(e) => setQty(it.id, e.target.value)} placeholder="0"
                  className={`w-14 text-center border border-stone-200 rounded-md px-1 py-1 text-sm focus:outline-none focus:ring-2 ${t.ring}`} />
                <button type="button" onClick={() => bump(it, 1)}
                  className="w-7 h-7 rounded-md border border-stone-300 flex items-center justify-center text-stone-500 hover:bg-stone-100 shrink-0">
                  <Plus className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          );
        })}
        {filtered.length === 0 && (
          <p className="col-span-full text-center text-sm text-stone-400 py-8">Sin resultados{search ? ` para "${search}"` : ""}.</p>
        )}
      </div>
    </div>
  );
}

function SelectionSummary({ allItems, quantities }) {
  const [open, setOpen] = useState(false);
  const selected = allItems.filter((it) => (parseInt(quantities[it.id], 10) || 0) > 0);
  const total = selected.reduce((s, it) => s + (parseInt(quantities[it.id], 10) || 0), 0);

  return (
    <div className="mt-6 border border-stone-200 rounded-lg overflow-hidden">
      <button type="button" onClick={() => setOpen((o) => !o)}
        className="w-full flex items-center justify-between px-4 py-3 bg-stone-50 hover:bg-stone-100 text-sm font-medium text-stone-700 transition-colors">
        <span>{selected.length} artículo{selected.length !== 1 ? "s" : ""} seleccionado{selected.length !== 1 ? "s" : ""} · {total} unidades en total</span>
        <ChevronDown className={`w-4 h-4 text-stone-400 transition-transform ${open ? "rotate-180" : ""}`} />
      </button>
      {open && (
        <div className="divide-y divide-stone-100 max-h-64 overflow-y-auto">
          {selected.length === 0 && <p className="px-4 py-4 text-sm text-stone-400 text-center">Todavía no has seleccionado artículos.</p>}
          {selected.map((it) => (
            <div key={it.id} className="flex items-center justify-between px-4 py-2 text-sm">
              <span className="text-stone-700">{it.item}</span>
              <span className="text-stone-500 font-medium">{quantities[it.id]} {it.unidad}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function FormView({ config, data, onGoAdmin }) {
  const hasActivePeriod = !!config?.activeFormat && !!config?.periodSlug;

  if (!hasActivePeriod) {
    return (
      <div className="max-w-lg mx-auto px-6 py-16 text-center">
        <AlertCircle className="w-8 h-8 mx-auto text-amber-600 mb-3" />
        <h2 className="text-lg font-semibold text-stone-800">No hay una solicitud activa</h2>
        <p className="text-sm text-stone-500 mt-2">
          El administrador todavía no ha abierto la solicitud de suministros de este bimestre. Intenta más tarde.
        </p>
        <button onClick={onGoAdmin} className="mt-8 text-xs text-stone-400 hover:text-stone-600 underline underline-offset-2">
          Acceso administración
        </button>
      </div>
    );
  }

  return config.activeFormat === "CARTIMEX"
    ? <CartimexForm config={config} areas={data.cartimexAreas} items={data.cartimexItems} onGoAdmin={onGoAdmin} />
    : <ComputronForm config={config} tiendas={data.computronTiendas} itemsOficina={data.computronItemsOficina} itemsLimpieza={data.computronItemsLimpieza} onGoAdmin={onGoAdmin} />;
}

function useExistingCheck(format, periodSlug, entityId) {
  const [checking, setChecking] = useState(false);
  const [existing, setExisting] = useState(null);
  useEffect(() => {
    if (!entityId) { setExisting(null); return; }
    (async () => {
      setChecking(true);
      const raw = await safeGet(submissionKey(format, periodSlug, entityId), true);
      setExisting(parseJSON(raw, null));
      setChecking(false);
    })();
  }, [format, periodSlug, entityId]);
  return { checking, existing };
}

function AlreadySubmittedBanner({ existing }) {
  return (
    <div className="mt-5 rounded-lg border border-amber-300 bg-amber-50 px-4 py-4">
      <div className="flex items-start gap-2">
        <AlertCircle className="w-4 h-4 text-amber-600 mt-0.5 shrink-0" />
        <div>
          <p className="text-sm font-medium text-amber-800">{existing.entityLabel} ya envió su solicitud para este periodo.</p>
          <p className="text-xs text-amber-700 mt-1">
            Enviada por {existing.nombre} {existing.apellido} el {new Date(existing.date).toLocaleString("es-EC")}.
            Solo se permite un envío por periodo. Si necesitas corregirla, contacta al administrador.
          </p>
        </div>
      </div>
    </div>
  );
}

function DoneScreen({ entityLabel, periodLabel, total }) {
  return (
    <div className="max-w-lg mx-auto px-6 py-16 text-center">
      <div className="w-14 h-14 mx-auto rounded-full bg-emerald-50 border-2 border-emerald-500 flex items-center justify-center mb-4">
        <CheckCircle2 className="w-7 h-7 text-emerald-600" />
      </div>
      <h2 className="text-lg font-semibold text-stone-800">Solicitud enviada</h2>
      <p className="text-sm text-stone-500 mt-2">
        La solicitud de <span className="font-medium text-stone-700">{entityLabel}</span> para{" "}
        <span className="font-medium text-stone-700">{periodLabel}</span> quedó registrada.
      </p>
      <p className="text-xs text-stone-400 mt-1">Total de unidades solicitadas: {total}</p>
    </div>
  );
}

function CategoryTabs({ active, onChange, countOficina, countLimpieza, theme }) {
  const t = theme || THEME.COMPUTRON;
  return (
    <div className="flex gap-1 border-b border-stone-200 mt-6">
      <button type="button" onClick={() => onChange("oficina")}
        className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${active === "oficina" ? `${t.tabActiveBorder} text-stone-800` : "border-transparent text-stone-400 hover:text-stone-600"}`}>
        Oficina
        {countOficina > 0 && <span className={`text-[10px] rounded-full px-1.5 py-0.5 font-semibold ${t.badgeBg} ${t.badgeText}`}>{countOficina}</span>}
      </button>
      <button type="button" onClick={() => onChange("limpieza")}
        className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${active === "limpieza" ? `${t.tabActiveBorder} text-stone-800` : "border-transparent text-stone-400 hover:text-stone-600"}`}>
        Limpieza
        {countLimpieza > 0 && <span className={`text-[10px] rounded-full px-1.5 py-0.5 font-semibold ${t.badgeBg} ${t.badgeText}`}>{countLimpieza}</span>}
      </button>
    </div>
  );
}

function ConfirmModal({ title, message, confirmLabel = "Confirmar y enviar", onConfirm, onCancel, loading, theme }) {
  const t = theme || THEME.CARTIMEX;
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4">
      <div className="bg-white rounded-xl shadow-xl max-w-sm w-full p-5">
        <h3 className="text-base font-semibold text-stone-800">{title}</h3>
        <div className="text-sm text-stone-600 mt-2 leading-relaxed">{message}</div>
        <div className="flex gap-2 mt-5">
          <button onClick={onCancel} disabled={loading}
            className="flex-1 border border-stone-300 text-stone-600 text-sm font-medium rounded-lg py-2.5 hover:bg-stone-50 disabled:opacity-50">
            Revisar de nuevo
          </button>
          <button onClick={onConfirm} disabled={loading}
            className={`flex-1 ${t.primary} ${t.primaryHover} disabled:opacity-60 text-white text-sm font-medium rounded-lg py-2.5 transition-colors`}>
            {loading ? "Enviando..." : confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

function FloatingSubmitBar({ theme, count, error, onSubmit, disabled }) {
  const t = theme || THEME.CARTIMEX;
  return (
    <div className="fixed bottom-0 inset-x-0 z-40 bg-white/95 backdrop-blur border-t border-stone-200 shadow-[0_-4px_16px_rgba(0,0,0,0.08)]">
      <div className="max-w-2xl mx-auto px-6 py-3 flex items-center justify-between gap-3">
        <div className="min-w-0">
          {error ? (
            <p className="text-xs text-red-600 font-medium truncate">{error}</p>
          ) : (
            <p className="text-xs text-stone-600">
              <span className="font-semibold text-stone-800">{count}</span> artículo{count !== 1 ? "s" : ""} seleccionado{count !== 1 ? "s" : ""}
            </p>
          )}
        </div>
        <button type="button" onClick={onSubmit} disabled={disabled}
          className={`shrink-0 ${t.primary} ${t.primaryHover} disabled:opacity-60 text-white font-medium text-sm rounded-lg px-5 py-2.5 transition-colors shadow-sm`}>
          Enviar solicitud
        </button>
      </div>
    </div>
  );
}

function FormHeader({ subtitle, periodLabel, stamp, theme }) {
  const t = theme || THEME.CARTIMEX;
  return (
    <div className="border-b-2 border-dashed border-stone-300 bg-white">
      <div className="px-6 pt-6 flex items-start justify-between">
        <div>
          <p className={`text-[11px] tracking-[0.2em] font-semibold ${t.subtitleText}`}>{subtitle}</p>
          <h1 className="text-xl font-bold text-stone-800 mt-0.5">Solicitud de Suministros</h1>
          <p className="text-xs text-stone-500 mt-1">Periodo: <span className="font-medium text-stone-700">{periodLabel}</span></p>
        </div>
        <Stamp theme={t}>{stamp}</Stamp>
      </div>
      <Perforation />
    </div>
  );
}

/* --------------------------- CARTIMEX form --------------------------- */

function CartimexForm({ config, areas, items, onGoAdmin }) {
  const [nombre, setNombre] = useState("");
  const [apellido, setApellido] = useState("");
  const [correo, setCorreo] = useState("");
  const [areaId, setAreaId] = useState("");
  const [quantities, setQuantities] = useState({});
  const [observaciones, setObservaciones] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [done, setDone] = useState(false);
  const [error, setError] = useState("");

  const { checking, existing } = useExistingCheck("CARTIMEX", config.periodSlug, areaId);

  const setQty = (id, val) => {
    const n = val === "" ? "" : Math.max(0, parseInt(val, 10) || 0);
    setQuantities((prev) => ({ ...prev, [id]: n }));
  };
  const total = useMemo(() => Object.values(quantities).reduce((a, b) => a + (parseInt(b, 10) || 0), 0), [quantities]);
  const selectedCount = useMemo(() => items.filter((it) => (parseInt(quantities[it.id], 10) || 0) > 0).length, [items, quantities]);
  const areaLabel = areas.find((a) => a.id === areaId)?.label || areaId;

  const infoComplete = nombre.trim() && apellido.trim() && EMAIL_RE.test(correo.trim());

  const validate = () => {
    setError("");
    if (!nombre.trim() || !apellido.trim()) { setError("Ingresa tu nombre y apellido."); return false; }
    if (!EMAIL_RE.test(correo.trim())) { setError("Ingresa un correo válido."); return false; }
    if (!areaId) { setError("Selecciona tu departamento."); return false; }
    const any = Object.values(quantities).some((v) => (parseInt(v, 10) || 0) > 0);
    if (!any) { setError("Ingresa al menos una cantidad mayor a 0."); return false; }
    return true;
  };

  const handleReview = () => { if (validate()) setShowConfirm(true); };

  const confirmSubmit = async () => {
    const clean = {};
    Object.entries(quantities).forEach(([k, v]) => { const n = parseInt(v, 10) || 0; if (n > 0) clean[k] = n; });

    setSubmitting(true);
    const record = {
      format: "CARTIMEX", period: config.periodSlug, periodLabel: config.periodLabel,
      entity: areaId, entityLabel: areaLabel,
      nombre: nombre.trim(), apellido: apellido.trim(), correo: correo.trim(),
      quantities: clean, observaciones: observaciones.trim(), date: new Date().toISOString(),
    };
    const ok = await safeSet(submissionKey("CARTIMEX", config.periodSlug, areaId), JSON.stringify(record), true);
    setSubmitting(false);
    setShowConfirm(false);
    if (ok) setDone(true); else setError("No se pudo guardar la solicitud. Intenta de nuevo.");
  };

  if (done) return <DoneScreen entityLabel={areaLabel} periodLabel={config.periodLabel} total={total} />;

  const theme = THEME.CARTIMEX;
  const showItems = !checking && areaId && !existing;

  return (
    <div className={`max-w-2xl mx-auto ${showItems ? "pb-28" : "pb-16"}`}>
      <FormHeader subtitle="CARTIMEX · OFICINA" periodLabel={config.periodLabel} stamp={<>SOLIC.<br/>INTERNA</>} theme={theme} />
      <div className="px-6 pt-6">
        <PersonalInfoFields nombre={nombre} apellido={apellido} correo={correo} setNombre={setNombre} setApellido={setApellido} setCorreo={setCorreo} theme={theme} />

        <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mt-5">Departamento</label>
        <select value={areaId} onChange={(e) => { setAreaId(e.target.value); setQuantities({}); setObservaciones(""); setError(""); }}
          disabled={!infoComplete}
          className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 ${theme.ring} disabled:bg-stone-100 disabled:text-stone-400`}>
          <option value="">{infoComplete ? "Selecciona tu departamento..." : "Completa tus datos primero"}</option>
          {areas.map((a) => <option key={a.id} value={a.id}>{a.label}</option>)}
        </select>

        {checking && <p className="text-xs text-stone-400 mt-3">Verificando...</p>}
        {!checking && areaId && existing && <AlreadySubmittedBanner existing={existing} />}

        {showItems && (
          <>
            <div className="mt-5">
              <ItemCatalog items={items} quantities={quantities} setQty={setQty} theme={theme} />
            </div>

            <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mt-5">Observaciones (opcional)</label>
            <textarea value={observaciones} onChange={(e) => setObservaciones(e.target.value)} rows={3}
              placeholder="Notas adicionales sobre esta solicitud..."
              className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 ${theme.ring} resize-none`} />

            <SelectionSummary allItems={items} quantities={quantities} />
          </>
        )}
      </div>
      <div className="text-center mt-10">
        <button onClick={onGoAdmin} className="text-xs text-stone-400 hover:text-stone-600 underline underline-offset-2">Acceso administración</button>
      </div>
      {showItems && (
        <FloatingSubmitBar theme={theme} count={selectedCount} error={error} onSubmit={handleReview} disabled={submitting} />
      )}
      {showConfirm && (
        <ConfirmModal
          title="Confirmar solicitud"
          message={<>Vas a enviar la solicitud de <b>{areaLabel}</b> con <b>{selectedCount}</b> artículo{selectedCount !== 1 ? "s" : ""} a nombre de <b>{nombre} {apellido}</b>. Una vez enviada no podrás modificarla. ¿Confirmas?</>}
          onConfirm={confirmSubmit}
          onCancel={() => setShowConfirm(false)}
          loading={submitting}
          theme={theme}
        />
      )}
    </div>
  );
}

/* --------------------------- COMPUTRON form --------------------------- */

function ComputronForm({ config, tiendas, itemsOficina, itemsLimpieza, onGoAdmin }) {
  const [nombre, setNombre] = useState("");
  const [apellido, setApellido] = useState("");
  const [correo, setCorreo] = useState("");
  const [tiendaId, setTiendaId] = useState("");
  const [quantities, setQuantities] = useState({});
  const [observaciones, setObservaciones] = useState("");
  const [category, setCategory] = useState("oficina");
  const [submitting, setSubmitting] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [done, setDone] = useState(false);
  const [error, setError] = useState("");

  const { checking, existing } = useExistingCheck("COMPUTRON", config.periodSlug, tiendaId);

  const setQty = (id, val) => {
    const n = val === "" ? "" : Math.max(0, parseInt(val, 10) || 0);
    setQuantities((prev) => ({ ...prev, [id]: n }));
  };
  const total = useMemo(() => Object.values(quantities).reduce((a, b) => a + (parseInt(b, 10) || 0), 0), [quantities]);
  const allItems = useMemo(() => [...itemsOficina, ...itemsLimpieza], [itemsOficina, itemsLimpieza]);
  const selectedCount = useMemo(() => allItems.filter((it) => (parseInt(quantities[it.id], 10) || 0) > 0).length, [allItems, quantities]);
  const tiendaLabel = tiendas.find((t) => t.id === tiendaId)?.label || tiendaId;
  const infoComplete = nombre.trim() && apellido.trim() && EMAIL_RE.test(correo.trim());

  const validate = () => {
    setError("");
    if (!nombre.trim() || !apellido.trim()) { setError("Ingresa tu nombre y apellido."); return false; }
    if (!EMAIL_RE.test(correo.trim())) { setError("Ingresa un correo válido."); return false; }
    if (!tiendaId) { setError("Selecciona tu tienda."); return false; }
    const any = Object.values(quantities).some((v) => (parseInt(v, 10) || 0) > 0);
    if (!any) { setError("Ingresa al menos una cantidad mayor a 0 (Oficina o Limpieza)."); return false; }
    return true;
  };

  const handleReview = () => { if (validate()) setShowConfirm(true); };

  const confirmSubmit = async () => {
    const clean = {};
    Object.entries(quantities).forEach(([k, v]) => { const n = parseInt(v, 10) || 0; if (n > 0) clean[k] = n; });

    setSubmitting(true);
    const record = {
      format: "COMPUTRON", period: config.periodSlug, periodLabel: config.periodLabel,
      entity: tiendaId, entityLabel: tiendaLabel,
      nombre: nombre.trim(), apellido: apellido.trim(), correo: correo.trim(),
      quantities: clean, observaciones: observaciones.trim(), date: new Date().toISOString(),
    };
    const ok = await safeSet(submissionKey("COMPUTRON", config.periodSlug, tiendaId), JSON.stringify(record), true);
    setSubmitting(false);
    setShowConfirm(false);
    if (ok) setDone(true); else setError("No se pudo guardar la solicitud. Intenta de nuevo.");
  };

  if (done) return <DoneScreen entityLabel={tiendaLabel} periodLabel={config.periodLabel} total={total} />;

  const theme = THEME.COMPUTRON;
  const showItems = !checking && tiendaId && !existing;

  return (
    <div className={`max-w-2xl mx-auto ${showItems ? "pb-28" : "pb-16"}`}>
      <FormHeader subtitle="COMPUTRON · TIENDAS" periodLabel={config.periodLabel} stamp={<>SOLIC.<br/>TIENDA</>} theme={theme} />
      <div className="px-6 pt-6">
        <PersonalInfoFields nombre={nombre} apellido={apellido} correo={correo} setNombre={setNombre} setApellido={setApellido} setCorreo={setCorreo} theme={theme} />

        <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mt-5">Tienda</label>
        <select value={tiendaId} onChange={(e) => { setTiendaId(e.target.value); setQuantities({}); setObservaciones(""); setError(""); }}
          disabled={!infoComplete}
          className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 ${theme.ring} disabled:bg-stone-100 disabled:text-stone-400`}>
          <option value="">{infoComplete ? "Selecciona tu tienda..." : "Completa tus datos primero"}</option>
          {tiendas.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
        </select>

        {checking && <p className="text-xs text-stone-400 mt-3">Verificando...</p>}
        {!checking && tiendaId && existing && <AlreadySubmittedBanner existing={existing} />}

        {showItems && (
          <>
            <CategoryTabs
              active={category}
              onChange={setCategory}
              countOficina={itemsOficina.filter((it) => (parseInt(quantities[it.id], 10) || 0) > 0).length}
              countLimpieza={itemsLimpieza.filter((it) => (parseInt(quantities[it.id], 10) || 0) > 0).length}
              theme={theme}
            />
            <div className="mt-4">
              {category === "oficina" ? (
                <ItemCatalog items={itemsOficina} quantities={quantities} setQty={setQty} theme={theme} />
              ) : (
                <ItemCatalog items={itemsLimpieza} quantities={quantities} setQty={setQty} theme={theme} />
              )}
            </div>

            <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mt-5">Observaciones (opcional)</label>
            <textarea value={observaciones} onChange={(e) => setObservaciones(e.target.value)} rows={3}
              placeholder="Notas adicionales sobre esta solicitud..."
              className={`mt-1.5 w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 ${theme.ring} resize-none`} />

            <SelectionSummary allItems={allItems} quantities={quantities} />
          </>
        )}
      </div>
      <div className="text-center mt-10">
        <button onClick={onGoAdmin} className="text-xs text-stone-400 hover:text-stone-600 underline underline-offset-2">Acceso administración</button>
      </div>
      {showItems && (
        <FloatingSubmitBar theme={theme} count={selectedCount} error={error} onSubmit={handleReview} disabled={submitting} />
      )}
      {showConfirm && (
        <ConfirmModal
          title="Confirmar solicitud"
          message={<>Vas a enviar la solicitud de <b>{tiendaLabel}</b> con <b>{selectedCount}</b> artículo{selectedCount !== 1 ? "s" : ""} a nombre de <b>{nombre} {apellido}</b>. Una vez enviada no podrás modificarla. ¿Confirmas?</>}
          onConfirm={confirmSubmit}
          onCancel={() => setShowConfirm(false)}
          loading={submitting}
          theme={theme}
        />
      )}
    </div>
  );
}

/* ==================================================================== */
/* ADMIN GATE                                                             */
/* ==================================================================== */

function AdminGate({ config, onBack, onSuccess }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const submit = () => { if (pw === config.adminPassword) onSuccess(); else setErr("Contraseña incorrecta."); };
  return (
    <div className="max-w-sm mx-auto px-6 py-20">
      <button onClick={onBack} className="flex items-center gap-1 text-xs text-stone-400 hover:text-stone-600 mb-8">
        <ArrowLeft className="w-3.5 h-3.5" /> Volver al formulario
      </button>
      <div className="flex flex-col items-center text-center">
        <div className="w-12 h-12 rounded-full bg-stone-800 flex items-center justify-center mb-4"><Lock className="w-5 h-5 text-white" /></div>
        <h2 className="text-base font-semibold text-stone-800">Acceso administración</h2>
        <p className="text-xs text-stone-400 mt-1 mb-6">Ingresa la contraseña para ver el informe consolidado.</p>
        <input type="password" value={pw} onChange={(e) => { setPw(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()}
          placeholder="Contraseña" autoFocus
          className="w-full border border-stone-300 rounded-lg px-3 py-2.5 text-sm text-center focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500" />
        {err && <p className="text-xs text-red-600 mt-2">{err}</p>}
        <button onClick={submit} className="mt-4 w-full bg-stone-800 hover:bg-stone-900 text-white text-sm font-medium rounded-lg py-2.5 transition-colors">Entrar</button>
      </div>
    </div>
  );
}

/* ==================================================================== */
/* ADMIN VIEW                                                             */
/* ==================================================================== */

function AdminView(props) {
  const { config, data, onExit } = props;
  const [tab, setTab] = useState("informe");

  return (
    <div className="max-w-6xl mx-auto pb-16">
      <div className="border-b border-stone-200 bg-white px-6 pt-5">
        <div className="flex items-center justify-between">
          <div>
            <p className="text-[11px] tracking-[0.2em] text-stone-400 font-medium">ADMINISTRACIÓN</p>
            <h1 className="text-lg font-bold text-stone-800 mt-0.5">Suministros CARTIMEX / COMPUTRON</h1>
          </div>
          <button onClick={onExit} className="flex items-center gap-1 text-xs text-stone-400 hover:text-stone-600"><ArrowLeft className="w-3.5 h-3.5" /> Salir</button>
        </div>
        <div className="flex gap-1 mt-5 flex-wrap">
          <TabButton icon={FileText} label="Informe" active={tab === "informe"} onClick={() => setTab("informe")} />
          <TabButton icon={Store} label="Tiendas / Departamentos" active={tab === "entidades"} onClick={() => setTab("entidades")} />
          <TabButton icon={Package} label="Artículos" active={tab === "articulos"} onClick={() => setTab("articulos")} />
          <TabButton icon={Settings} label="Configuración" active={tab === "config"} onClick={() => setTab("config")} />
        </div>
      </div>
      <div className="px-6 pt-6">
        {tab === "informe" && <InformeTab config={config} data={data} />}
        {tab === "entidades" && <EntidadesTab data={data} onCartimexAreasChange={props.onCartimexAreasChange} onComputronTiendasChange={props.onComputronTiendasChange} />}
        {tab === "articulos" && <ArticulosTab data={data} onCartimexItemsChange={props.onCartimexItemsChange} onComputronItemsOficinaChange={props.onComputronItemsOficinaChange} onComputronItemsLimpiezaChange={props.onComputronItemsLimpiezaChange} />}
        {tab === "config" && <ConfigTab config={config} data={data} onConfigChange={props.onConfigChange} />}
      </div>
    </div>
  );
}

function TabButton({ icon: Icon, label, active, onClick }) {
  return (
    <button onClick={onClick}
      className={`flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium rounded-t-lg border-b-2 transition-colors ${active ? "border-amber-600 text-stone-800" : "border-transparent text-stone-400 hover:text-stone-600"}`}>
      <Icon className="w-4 h-4" /> {label}
    </button>
  );
}

function FormatSwitch({ value, onChange }) {
  return (
    <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden text-sm">
      <button onClick={() => onChange("CARTIMEX")} className={`px-3.5 py-1.5 font-medium flex items-center gap-1.5 ${value === "CARTIMEX" ? "bg-stone-800 text-white" : "bg-white text-stone-600 hover:bg-stone-50"}`}>
        <Building2 className="w-3.5 h-3.5" /> CARTIMEX
      </button>
      <button onClick={() => onChange("COMPUTRON")} className={`px-3.5 py-1.5 font-medium flex items-center gap-1.5 ${value === "COMPUTRON" ? "bg-stone-800 text-white" : "bg-white text-stone-600 hover:bg-stone-50"}`}>
        <Store className="w-3.5 h-3.5" /> COMPUTRON
      </button>
    </div>
  );
}

/* ---------------------------- Informe ------------------------------ */

function InformeTab({ config, data }) {
  const [format, setFormat] = useState(config.activeFormat || "CARTIMEX");
  const [periods, setPeriods] = useState([]);
  const [selectedPeriod, setSelectedPeriod] = useState("");
  const [submissions, setSubmissions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filter, setFilter] = useState("todos"); // todos | enviaron | pendientes
  const [expandedId, setExpandedId] = useState(null);
  const [search, setSearch] = useState("");

  useEffect(() => {
    (async () => {
      const list = parseJSON(await safeGet(PERIODS_KEY, true), []);
      setPeriods(list);
    })();
  }, []);

  const periodsForFormat = periods.filter((p) => p.format === format);

  useEffect(() => {
    if (periodsForFormat.length && !periodsForFormat.find((p) => p.slug === selectedPeriod)) {
      setSelectedPeriod(periodsForFormat[periodsForFormat.length - 1].slug);
    }
    if (!periodsForFormat.length) setSelectedPeriod("");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [format, periods]);

  useEffect(() => {
    if (!selectedPeriod) { setSubmissions([]); setLoading(false); return; }
    (async () => {
      setLoading(true);
      const keys = await safeList(`submission:${format}:${selectedPeriod}:`, true);
      const recs = [];
      for (const k of keys) { const raw = await safeGet(k, true); const rec = parseJSON(raw, null); if (rec) recs.push(rec); }
      setSubmissions(recs);
      setLoading(false);
      setExpandedId(null);
    })();
  }, [format, selectedPeriod]);

  const entities = format === "CARTIMEX" ? data.cartimexAreas : data.computronTiendas;

  const itemMap = useMemo(() => {
    const m = {};
    const all = format === "CARTIMEX" ? data.cartimexItems : [...data.computronItemsOficina, ...data.computronItemsLimpieza];
    all.forEach((it) => { m[it.id] = it; });
    return m;
  }, [format, data]);

  const entityRows = useMemo(() => {
    return entities.map((en) => {
      const sub = submissions.find((s) => s.entity === en.id);
      let totalOficina = 0, totalLimpieza = 0, totalGeneral = 0;
      if (sub) {
        Object.entries(sub.quantities).forEach(([id, qty]) => {
          totalGeneral += qty;
          if (format === "COMPUTRON") {
            if (id.startsWith("ofc")) totalOficina += qty;
            else if (id.startsWith("lmp")) totalLimpieza += qty;
          }
        });
      }
      return { ...en, submitted: !!sub, sub, totalOficina, totalLimpieza, totalGeneral };
    });
  }, [entities, submissions, format]);

  const submittedCount = entityRows.filter((r) => r.submitted).length;

  const filteredRows = entityRows.filter((r) => {
    if (filter === "enviaron" && !r.submitted) return false;
    if (filter === "pendientes" && r.submitted) return false;
    if (search.trim() && !r.label.toLowerCase().includes(search.trim().toLowerCase())) return false;
    return true;
  });

  const detailLists = (sub) => {
    const entries = Object.entries(sub.quantities)
      .map(([id, qty]) => ({ ...(itemMap[id] || {}), id, qty }))
      .filter((x) => x.item);
    if (format !== "COMPUTRON") return { general: entries.sort((a, b) => a.numero - b.numero) };
    return {
      oficina: entries.filter((x) => x.id.startsWith("ofc")).sort((a, b) => a.numero - b.numero),
      limpieza: entries.filter((x) => x.id.startsWith("lmp")).sort((a, b) => a.numero - b.numero),
    };
  };

  /* ---- export a Excel (matriz completa, igual al formato original) ---- */
  const buildSheetAOA = (rowsForSheet, entitiesForSheet, totalsForSheet, title) => {
    const aoa = [];
    aoa.push([]); aoa.push(["SUMINISTROS"]); aoa.push([title]); aoa.push([]); aoa.push([]);
    aoa.push(["NÚMERO", "PROVEEDOR", "ITEM", "UNIDAD DE MEDIDA", ...entitiesForSheet.map((e) => e.label.toUpperCase()), "TOTAL SOLICITADO EN EL PERIODO", "OBSERVACIONES"]);
    rowsForSheet.forEach((r) => {
      aoa.push([r.numero, r.proveedor, r.item, r.unidad, ...entitiesForSheet.map((e) => r.byEntity[e.id] || 0), r.total, r.obs]);
    });
    aoa.push(["", "", "TOTAL", "", ...entitiesForSheet.map((e) => totalsForSheet[e.id]), totalsForSheet.TOTAL, ""]);
    return aoa;
  };

  const buildItemRows = (its) => its.map((it) => {
    const byEntity = {}; let total = 0; const obs = [];
    entities.forEach((en) => {
      const sub = submissions.find((s) => s.entity === en.id);
      const qty = sub?.quantities?.[it.id] || 0;
      byEntity[en.id] = qty; total += qty;
      if (sub?.observaciones && qty > 0) obs.push(`${en.label}: ${sub.observaciones}`);
    });
    return { ...it, byEntity, total, obs: obs.join(" · ") };
  });
  const buildTotals = (rws) => {
    const t = {}; entities.forEach((en) => { t[en.id] = rws.reduce((s, r) => s + (r.byEntity[en.id] || 0), 0); });
    t.TOTAL = rws.reduce((s, r) => s + r.total, 0); return t;
  };

  const exportExcel = async () => {
    const periodLabel = periodsForFormat.find((p) => p.slug === selectedPeriod)?.label || selectedPeriod;
    const wb = XLSX.utils.book_new();

    if (format === "CARTIMEX") {
      const rows = buildItemRows(data.cartimexItems);
      const aoa = buildSheetAOA(rows, entities, buildTotals(rows), `OFICINA · ${periodLabel}`);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 8 }, { wch: 12 }, { wch: 40 }, { wch: 10 }, ...entities.map(() => ({ wch: 11 })), { wch: 14 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, ws, "Consolidado");
    } else {
      const rowsOfc = buildItemRows(data.computronItemsOficina);
      const rowsLmp = buildItemRows(data.computronItemsLimpieza);

      const aoaOfc = buildSheetAOA(rowsOfc, entities, buildTotals(rowsOfc), `OFICINA · ${periodLabel}`);
      const wsOfc = XLSX.utils.aoa_to_sheet(aoaOfc);
      wsOfc["!cols"] = [{ wch: 8 }, { wch: 12 }, { wch: 40 }, { wch: 10 }, ...entities.map(() => ({ wch: 10 })), { wch: 14 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, wsOfc, "Suministros Oficina");

      const aoaLmp = buildSheetAOA(rowsLmp, entities, buildTotals(rowsLmp), `LIMPIEZA · ${periodLabel}`);
      const wsLmp = XLSX.utils.aoa_to_sheet(aoaLmp);
      wsLmp["!cols"] = [{ wch: 8 }, { wch: 12 }, { wch: 40 }, { wch: 10 }, ...entities.map(() => ({ wch: 10 })), { wch: 14 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, wsLmp, "Suministros Limpieza");
    }

    const detalle = [["ÁREA/TIENDA", "NOMBRE", "APELLIDO", "CORREO", "FECHA DE ENVÍO"]];
    submissions.forEach((s) => detalle.push([s.entityLabel, s.nombre, s.apellido, s.correo, new Date(s.date).toLocaleString("es-EC")]));
    const wsDet = XLSX.utils.aoa_to_sheet(detalle);
    wsDet["!cols"] = [{ wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 28 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, wsDet, "Detalle de envíos");

    XLSX.writeFile(wb, `Consolidado_${format}_${selectedPeriod || "periodo"}.xlsx`);
  };

  const noun = format === "CARTIMEX" ? "departamentos" : "tiendas";

  return (
    <div>
      <div className="flex flex-wrap items-end justify-between gap-3 mb-4">
        <div className="flex flex-wrap items-end gap-3">
          <div>
            <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mb-1">Marca</label>
            <FormatSwitch value={format} onChange={setFormat} />
          </div>
          <div>
            <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide">Periodo</label>
            <select value={selectedPeriod} onChange={(e) => setSelectedPeriod(e.target.value)}
              className="mt-1 block border border-stone-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-amber-500">
              {periodsForFormat.length === 0 && <option value="">Sin periodos registrados</option>}
              {periodsForFormat.map((p) => <option key={p.slug} value={p.slug}>{p.label}</option>)}
            </select>
          </div>
        </div>
        <button onClick={exportExcel} disabled={!selectedPeriod || entities.length === 0}
          className="flex items-center gap-1.5 bg-stone-800 hover:bg-stone-900 disabled:opacity-40 text-white text-sm font-medium rounded-lg px-4 py-2.5 transition-colors">
          <Download className="w-4 h-4" /> Exportar a Excel
        </button>
      </div>

      {loading ? (
        <p className="text-sm text-stone-400 py-10 text-center">Cargando solicitudes...</p>
      ) : !selectedPeriod ? (
        <p className="text-sm text-stone-400 py-10 text-center">Todavía no hay un periodo de {format} registrado. Ve a Configuración.</p>
      ) : (
        <>
          <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
            <div className="flex gap-1.5">
              <button onClick={() => setFilter("todos")} className={`text-xs font-medium px-3 py-1.5 rounded-full border transition-colors ${filter === "todos" ? "bg-stone-800 border-stone-800 text-white" : "bg-white border-stone-300 text-stone-600 hover:bg-stone-50"}`}>
                Todos ({entityRows.length})
              </button>
              <button onClick={() => setFilter("enviaron")} className={`text-xs font-medium px-3 py-1.5 rounded-full border transition-colors ${filter === "enviaron" ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-stone-300 text-stone-600 hover:bg-stone-50"}`}>
                Enviaron ({submittedCount})
              </button>
              <button onClick={() => setFilter("pendientes")} className={`text-xs font-medium px-3 py-1.5 rounded-full border transition-colors ${filter === "pendientes" ? "bg-amber-600 border-amber-600 text-white" : "bg-white border-stone-300 text-stone-600 hover:bg-stone-50"}`}>
                Pendientes ({entityRows.length - submittedCount})
              </button>
            </div>
            <div className="relative">
              <Search className="w-4 h-4 text-stone-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder={`Buscar ${noun}...`}
                className="border border-stone-300 rounded-lg pl-9 pr-3 py-1.5 text-sm w-56 focus:outline-none focus:ring-2 focus:ring-amber-500" />
            </div>
          </div>

          <div className="border border-stone-200 rounded-lg divide-y divide-stone-100">
            {filteredRows.map((r) => {
              const expanded = expandedId === r.id;
              const lists = r.submitted ? detailLists(r.sub) : null;
              return (
                <div key={r.id}>
                  <button onClick={() => r.submitted && setExpandedId(expanded ? null : r.id)}
                    className={`w-full flex items-center justify-between gap-3 px-4 py-3 text-left transition-colors ${r.submitted ? "hover:bg-stone-50 cursor-pointer" : "cursor-default"}`}>
                    <div className="flex items-center gap-2.5 min-w-0">
                      <span className={`w-2 h-2 rounded-full shrink-0 ${r.submitted ? "bg-emerald-500" : "bg-stone-300"}`} />
                      <span className="text-sm font-medium text-stone-800 truncate">{r.label}</span>
                    </div>
                    <div className="flex items-center gap-4 shrink-0">
                      {format === "COMPUTRON" ? (
                        <div className="hidden sm:flex items-center gap-3 text-xs text-stone-500">
                          <span>Oficina: <b className="text-stone-700">{r.totalOficina}</b></span>
                          <span>Limpieza: <b className="text-stone-700">{r.totalLimpieza}</b></span>
                        </div>
                      ) : (
                        <span className="hidden sm:inline text-xs text-stone-500">Total: <b className="text-stone-700">{r.totalGeneral}</b></span>
                      )}
                      {r.submitted ? (
                        <span className="text-[11px] font-medium bg-emerald-50 text-emerald-700 rounded-full px-2.5 py-1 whitespace-nowrap">Enviado</span>
                      ) : (
                        <span className="text-[11px] font-medium bg-stone-100 text-stone-400 rounded-full px-2.5 py-1 whitespace-nowrap">Pendiente</span>
                      )}
                      {r.submitted && <ChevronDown className={`w-4 h-4 text-stone-400 transition-transform ${expanded ? "rotate-180" : ""}`} />}
                    </div>
                  </button>

                  {expanded && r.submitted && (
                    <div className="px-4 pb-4">
                      <p className="text-xs text-stone-400 mb-3">
                        Enviado por <span className="text-stone-600 font-medium">{r.sub.nombre} {r.sub.apellido}</span> ({r.sub.correo}) el {new Date(r.sub.date).toLocaleString("es-EC")}
                      </p>
                      {format === "COMPUTRON" ? (
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div>
                            <p className="text-xs font-semibold text-stone-600 mb-1.5">Oficina ({r.totalOficina} unidades)</p>
                            <div className="border border-stone-200 rounded-lg divide-y divide-stone-100 max-h-56 overflow-y-auto">
                              {lists.oficina.length === 0 && <p className="px-3 py-3 text-xs text-stone-400">Sin artículos.</p>}
                              {lists.oficina.map((it) => (
                                <div key={it.id} className="flex items-center justify-between px-3 py-1.5 text-xs">
                                  <span className="text-stone-700">{it.item}</span>
                                  <span className="text-stone-500 font-medium shrink-0 ml-2">{it.qty} {it.unidad}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                          <div>
                            <p className="text-xs font-semibold text-stone-600 mb-1.5">Limpieza ({r.totalLimpieza} unidades)</p>
                            <div className="border border-stone-200 rounded-lg divide-y divide-stone-100 max-h-56 overflow-y-auto">
                              {lists.limpieza.length === 0 && <p className="px-3 py-3 text-xs text-stone-400">Sin artículos.</p>}
                              {lists.limpieza.map((it) => (
                                <div key={it.id} className="flex items-center justify-between px-3 py-1.5 text-xs">
                                  <span className="text-stone-700">{it.item}</span>
                                  <span className="text-stone-500 font-medium shrink-0 ml-2">{it.qty} {it.unidad}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div className="border border-stone-200 rounded-lg divide-y divide-stone-100 max-h-72 overflow-y-auto">
                          {lists.general.map((it) => (
                            <div key={it.id} className="flex items-center justify-between px-3 py-1.5 text-xs">
                              <span className="text-stone-700">{it.item}</span>
                              <span className="text-stone-500 font-medium shrink-0 ml-2">{it.qty} {it.unidad}</span>
                            </div>
                          ))}
                        </div>
                      )}
                      {r.sub.observaciones && (
                        <p className="text-xs text-stone-500 mt-3"><span className="font-semibold text-stone-600">Observaciones:</span> {r.sub.observaciones}</p>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
            {filteredRows.length === 0 && <p className="px-4 py-8 text-center text-sm text-stone-400">Sin resultados.</p>}
          </div>
        </>
      )}
    </div>
  );
}

/* --------------------------- Entidades (tiendas/departamentos) ------------------------------ */

function EntidadesTab({ data, onCartimexAreasChange, onComputronTiendasChange }) {
  const [format, setFormat] = useState("CARTIMEX");
  const entities = format === "CARTIMEX" ? data.cartimexAreas : data.computronTiendas;
  const onChange = format === "CARTIMEX" ? onCartimexAreasChange : onComputronTiendasChange;
  const noun = format === "CARTIMEX" ? "departamento" : "tienda";

  const [newLabel, setNewLabel] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [draftLabel, setDraftLabel] = useState("");

  const add = () => {
    if (!newLabel.trim()) return;
    const id = makeId(entities.map((e) => e.id), newLabel);
    onChange([...entities, { id, label: newLabel.trim() }]);
    setNewLabel("");
  };
  const remove = (id) => onChange(entities.filter((e) => e.id !== id));
  const startEdit = (e) => { setEditingId(e.id); setDraftLabel(e.label); };
  const saveEdit = (id) => { onChange(entities.map((e) => (e.id === id ? { ...e, label: draftLabel } : e))); setEditingId(null); };

  return (
    <div className="max-w-2xl">
      <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mb-1">Marca</label>
      <FormatSwitch value={format} onChange={setFormat} />
      <p className="text-xs text-stone-500 mt-3 mb-4">
        {format === "CARTIMEX" ? "Departamentos que aparecen en el formulario y el informe de CARTIMEX." : "Tiendas que aparecen en el formulario y el informe de COMPUTRON."}
      </p>

      <div className="flex gap-2 mb-4">
        <input value={newLabel} onChange={(e) => setNewLabel(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()}
          placeholder={`Nombre de ${noun}...`} className="flex-1 border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" />
        <button onClick={add} className="flex items-center gap-1.5 bg-amber-600 hover:bg-amber-700 text-white text-sm font-medium rounded-lg px-3.5 py-2"><Plus className="w-4 h-4" /> Agregar</button>
      </div>

      <div className="border border-stone-200 rounded-lg divide-y divide-stone-100 max-h-[480px] overflow-y-auto">
        {entities.map((e) => (
          <div key={e.id} className="flex items-center justify-between px-3 py-2 text-sm">
            {editingId === e.id ? (
              <>
                <input value={draftLabel} onChange={(ev) => setDraftLabel(ev.target.value)} className="flex-1 border border-stone-300 rounded-md px-2 py-1 text-sm mr-2" />
                <div className="flex gap-2">
                  <button onClick={() => saveEdit(e.id)} className="text-emerald-600 hover:text-emerald-800"><CheckCircle2 className="w-4 h-4" /></button>
                  <button onClick={() => setEditingId(null)} className="text-stone-400 hover:text-stone-600"><X className="w-4 h-4" /></button>
                </div>
              </>
            ) : (
              <>
                <span className="text-stone-800">{e.label}</span>
                <div className="flex gap-3">
                  <button onClick={() => startEdit(e)} className="text-stone-400 hover:text-stone-700"><Pencil className="w-4 h-4" /></button>
                  <button onClick={() => remove(e.id)} className="text-stone-400 hover:text-red-600"><Trash2 className="w-4 h-4" /></button>
                </div>
              </>
            )}
          </div>
        ))}
        {entities.length === 0 && <p className="px-3 py-6 text-center text-sm text-stone-400">Sin {noun}s registrados.</p>}
      </div>
    </div>
  );
}

/* --------------------------- Artículos ------------------------------ */

function ArticulosTab({ data, onCartimexItemsChange, onComputronItemsOficinaChange, onComputronItemsLimpiezaChange }) {
  const [format, setFormat] = useState("CARTIMEX");
  const [category, setCategory] = useState("oficina");

  let items, onChange, idPrefix, editingLabel;
  if (format === "CARTIMEX") { items = data.cartimexItems; onChange = onCartimexItemsChange; idPrefix = "itm"; editingLabel = "CARTIMEX"; }
  else if (category === "oficina") { items = data.computronItemsOficina; onChange = onComputronItemsOficinaChange; idPrefix = "ofc"; editingLabel = "COMPUTRON · Oficina"; }
  else { items = data.computronItemsLimpieza; onChange = onComputronItemsLimpiezaChange; idPrefix = "lmp"; editingLabel = "COMPUTRON · Limpieza"; }

  return (
    <div>
      <div className="flex items-center gap-2 mb-4 bg-amber-50 border border-amber-200 rounded-lg px-3.5 py-2.5 w-fit">
        <Pencil className="w-3.5 h-3.5 text-amber-700 shrink-0" />
        <p className="text-xs text-amber-800">Estás editando: <span className="font-semibold">{editingLabel}</span> — los artículos que agregues se guardarán aquí.</p>
      </div>
      <div className="flex flex-wrap items-end gap-4 mb-4">
        <div>
          <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mb-1">Marca</label>
          <FormatSwitch value={format} onChange={setFormat} />
        </div>
        {format === "COMPUTRON" && (
          <div>
            <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mb-1">Categoría</label>
            <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden text-sm">
              <button onClick={() => setCategory("oficina")} className={`px-3 py-1.5 font-medium ${category === "oficina" ? "bg-amber-600 text-white" : "bg-white text-stone-600"}`}>Oficina</button>
              <button onClick={() => setCategory("limpieza")} className={`px-3 py-1.5 font-medium ${category === "limpieza" ? "bg-amber-600 text-white" : "bg-white text-stone-600"}`}>Limpieza</button>
            </div>
          </div>
        )}
      </div>
      <ItemsManager key={`${format}-${category}`} items={items} onChange={onChange} idPrefix={idPrefix} />
    </div>
  );
}

function ItemsManager({ items, onChange, idPrefix }) {
  const [search, setSearch] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({ proveedor: "", item: "", unidad: "" });
  const [newItem, setNewItem] = useState({ proveedor: "", item: "", unidad: "UND" });
  const [showNew, setShowNew] = useState(false);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return items;
    return items.filter((it) => it.item.toLowerCase().includes(q) || it.proveedor.toLowerCase().includes(q));
  }, [items, search]);

  const startEdit = (it) => { setEditingId(it.id); setDraft({ proveedor: it.proveedor, item: it.item, unidad: it.unidad }); };
  const saveEdit = (id) => { onChange(items.map((it) => (it.id === id ? { ...it, ...draft } : it))); setEditingId(null); };
  const remove = (id) => onChange(items.filter((it) => it.id !== id));
  const addItem = () => {
    if (!newItem.item.trim()) return;
    const nextNumero = items.length ? Math.max(...items.map((i) => i.numero)) + 1 : 1;
    const id = makeId(items.map((i) => i.id), `${idPrefix}${nextNumero}`);
    onChange([...items, { id, numero: nextNumero, ...newItem }]);
    setNewItem({ proveedor: "", item: "", unidad: "UND" });
    setShowNew(false);
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div className="relative">
          <Search className="w-4 h-4 text-stone-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar artículo..." className="border border-stone-300 rounded-lg pl-9 pr-3 py-2 text-sm w-64 focus:outline-none focus:ring-2 focus:ring-amber-500" />
        </div>
        <button onClick={() => setShowNew((s) => !s)} className="flex items-center gap-1.5 bg-amber-600 hover:bg-amber-700 text-white text-sm font-medium rounded-lg px-3.5 py-2"><Plus className="w-4 h-4" /> Nuevo artículo</button>
      </div>

      {showNew && (
        <div className="mb-4 border border-amber-200 bg-amber-50 rounded-lg p-4 grid grid-cols-1 sm:grid-cols-[1fr_2fr_5rem_auto] gap-2 items-center">
          <input value={newItem.proveedor} onChange={(e) => setNewItem({ ...newItem, proveedor: e.target.value })} placeholder="Proveedor" className="border border-stone-300 rounded-md px-2 py-1.5 text-sm" />
          <input value={newItem.item} onChange={(e) => setNewItem({ ...newItem, item: e.target.value })} placeholder="Nombre del artículo" className="border border-stone-300 rounded-md px-2 py-1.5 text-sm" />
          <input value={newItem.unidad} onChange={(e) => setNewItem({ ...newItem, unidad: e.target.value })} placeholder="Unidad" className="border border-stone-300 rounded-md px-2 py-1.5 text-sm" />
          <button onClick={addItem} className="bg-stone-800 hover:bg-stone-900 text-white text-sm rounded-md px-3 py-1.5">Agregar</button>
        </div>
      )}

      <div className="border border-stone-200 rounded-lg overflow-hidden">
        <div className="grid grid-cols-[2.5rem_1fr_1fr_5rem_5rem] bg-stone-100 text-[11px] font-semibold text-stone-500 uppercase tracking-wide px-3 py-2">
          <span>#</span><span>Proveedor</span><span>Artículo</span><span>Unidad</span><span></span>
        </div>
        <div className="divide-y divide-stone-100 max-h-[500px] overflow-y-auto">
          {filtered.map((it) => (
            <div key={it.id} className="grid grid-cols-[2.5rem_1fr_1fr_5rem_5rem] items-center px-3 py-2 text-sm">
              <span className="text-stone-400 font-mono text-xs">{it.numero}</span>
              {editingId === it.id ? (
                <>
                  <input value={draft.proveedor} onChange={(e) => setDraft({ ...draft, proveedor: e.target.value })} className="border border-stone-300 rounded-md px-2 py-1 text-xs mr-2" />
                  <input value={draft.item} onChange={(e) => setDraft({ ...draft, item: e.target.value })} className="border border-stone-300 rounded-md px-2 py-1 text-xs mr-2" />
                  <input value={draft.unidad} onChange={(e) => setDraft({ ...draft, unidad: e.target.value })} className="border border-stone-300 rounded-md px-2 py-1 text-xs" />
                  <div className="flex gap-2 justify-end">
                    <button onClick={() => saveEdit(it.id)} className="text-emerald-600 hover:text-emerald-800"><CheckCircle2 className="w-4 h-4" /></button>
                    <button onClick={() => setEditingId(null)} className="text-stone-400 hover:text-stone-600"><X className="w-4 h-4" /></button>
                  </div>
                </>
              ) : (
                <>
                  <span className="text-stone-500">{it.proveedor}</span>
                  <span className="text-stone-800">{it.item}</span>
                  <span className="text-stone-500">{it.unidad}</span>
                  <div className="flex gap-3 justify-end">
                    <button onClick={() => startEdit(it)} className="text-stone-400 hover:text-stone-700"><Pencil className="w-4 h-4" /></button>
                    <button onClick={() => remove(it.id)} className="text-stone-400 hover:text-red-600"><Trash2 className="w-4 h-4" /></button>
                  </div>
                </>
              )}
            </div>
          ))}
          {filtered.length === 0 && <p className="px-3 py-6 text-center text-sm text-stone-400">Sin artículos.</p>}
        </div>
      </div>
    </div>
  );
}

/* --------------------------- Configuración --------------------------- */

function ConfigTab({ config, data, onConfigChange }) {
  const [newFormat, setNewFormat] = useState("CARTIMEX");
  const [periodLabel, setPeriodLabel] = useState("");
  const [pw1, setPw1] = useState("");
  const [pw2, setPw2] = useState("");
  const [msg, setMsg] = useState("");
  const [respondedCount, setRespondedCount] = useState(null);
  const [periodsList, setPeriodsList] = useState([]);

  const fetchPeriods = useCallback(async () => {
    const list = parseJSON(await safeGet(PERIODS_KEY, true), []);
    setPeriodsList(list);
  }, []);

  useEffect(() => { fetchPeriods(); }, [fetchPeriods]);

  useEffect(() => {
    if (!config.activeFormat || !config.periodSlug) { setRespondedCount(null); return; }
    (async () => {
      const keys = await safeList(`submission:${config.activeFormat}:${config.periodSlug}:`, true);
      setRespondedCount(keys.length);
    })();
  }, [config.activeFormat, config.periodSlug]);

  const activeEntitiesCount = config.activeFormat === "CARTIMEX" ? data.cartimexAreas.length
    : config.activeFormat === "COMPUTRON" ? data.computronTiendas.length : null;

  const openNewPeriod = async () => {
    if (!periodLabel.trim()) return;
    const slug = slugify(periodLabel) || `PERIODO-${Date.now()}`;
    const next = { ...config, activeFormat: newFormat, periodSlug: slug, periodLabel: periodLabel.trim() };
    await onConfigChange(next);

    const list = parseJSON(await safeGet(PERIODS_KEY, true), []);
    if (!list.find((p) => p.format === newFormat && p.slug === slug)) {
      list.push({ format: newFormat, slug, label: periodLabel.trim() });
      await safeSet(PERIODS_KEY, JSON.stringify(list), true);
      setPeriodsList(list);
    }
    setPeriodLabel("");
    setMsg(`Periodo de ${newFormat} abierto. El formulario ya muestra ese formato a quien abra el link.`);
  };

  const closePeriod = async () => {
    await onConfigChange({ ...config, activeFormat: "", periodSlug: "", periodLabel: "" });
    setMsg("Periodo cerrado. Nadie podrá enviar solicitudes hasta que abras uno nuevo o reabras uno anterior.");
  };

  const reopenPeriod = async (p) => {
    await onConfigChange({ ...config, activeFormat: p.format, periodSlug: p.slug, periodLabel: p.label });
    setMsg(`Periodo "${p.label}" reabierto. Los jefes que ya habían enviado en ese periodo seguirán marcados como enviados; los que falten podrán completar su solicitud.`);
  };

  const changePassword = async () => {
    if (!pw1 || pw1.length < 4) { setMsg("La contraseña debe tener al menos 4 caracteres."); return; }
    if (pw1 !== pw2) { setMsg("Las contraseñas no coinciden."); return; }
    await onConfigChange({ ...config, adminPassword: pw1 });
    setPw1(""); setPw2("");
    setMsg("Contraseña actualizada.");
  };

  const reopenCandidates = periodsList.filter(
    (p) => p.format === newFormat && !(config.activeFormat === p.format && config.periodSlug === p.slug)
  );

  return (
    <div className="max-w-xl space-y-8">
      <div>
        <h3 className="text-sm font-semibold text-stone-800 flex items-center gap-1.5"><ClipboardList className="w-4 h-4" /> Periodo del formulario</h3>

        <div className={`mt-2 rounded-lg border px-3 py-2.5 text-sm ${config.activeFormat ? "border-emerald-200 bg-emerald-50 text-emerald-800" : "border-stone-200 bg-stone-50 text-stone-500"}`}>
          {config.activeFormat
            ? <>Abierto ahora: <span className="font-semibold">{config.activeFormat}</span> · <span className="font-semibold">{config.periodLabel}</span>{respondedCount !== null && <> · {respondedCount} de {activeEntitiesCount} han enviado</>}</>
            : "No hay ningún periodo abierto. El link no dejará enviar solicitudes."}
        </div>

        {config.activeFormat && (
          <button onClick={closePeriod}
            className="mt-3 flex items-center gap-1.5 border border-red-200 text-red-600 hover:bg-red-50 text-sm font-medium rounded-lg px-3.5 py-2 transition-colors">
            <StopCircle className="w-4 h-4" /> Cerrar periodo actual
          </button>
        )}

        <div className="mt-6">
          <label className="text-xs font-semibold text-stone-500 uppercase tracking-wide block mb-1">Marca</label>
          <FormatSwitch value={newFormat} onChange={setNewFormat} />
        </div>

        <p className="text-xs font-semibold text-stone-600 mt-5">Abrir un periodo nuevo</p>
        <div className="flex gap-2 mt-1.5">
          <input value={periodLabel} onChange={(e) => setPeriodLabel(e.target.value)} placeholder="Ej: Agosto 2026 - Bimestre 2"
            className="flex-1 border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" />
          <button onClick={openNewPeriod}
            className="flex items-center gap-1.5 bg-amber-600 hover:bg-amber-700 text-white text-sm font-medium rounded-lg px-4 py-2 whitespace-nowrap transition-colors">
            <PlayCircle className="w-4 h-4" /> Abrir periodo
          </button>
        </div>

        {reopenCandidates.length > 0 && (
          <div className="mt-5">
            <p className="text-xs font-semibold text-stone-600 mb-1.5">Reabrir un periodo anterior de {newFormat}</p>
            <p className="text-[11px] text-stone-400 mb-2">Útil para casos excepcionales, por ejemplo si una tienda necesita corregir o completar su envío.</p>
            <div className="border border-stone-200 rounded-lg divide-y divide-stone-100">
              {reopenCandidates.map((p) => (
                <div key={p.slug} className="flex items-center justify-between px-3 py-2 text-sm">
                  <span className="text-stone-700">{p.label}</span>
                  <button onClick={() => reopenPeriod(p)}
                    className="flex items-center gap-1 text-xs font-medium text-amber-700 hover:text-amber-800">
                    <RotateCcw className="w-3.5 h-3.5" /> Reabrir
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        <p className="text-[11px] text-stone-400 mt-4">Solo un formato puede estar activo a la vez. Los datos ya enviados nunca se pierden al cerrar o cambiar de periodo.</p>
      </div>

      <div>
        <h3 className="text-sm font-semibold text-stone-800 flex items-center gap-1.5"><Lock className="w-4 h-4" /> Contraseña de administración</h3>
        <div className="flex flex-col gap-2 mt-3 max-w-xs">
          <input type="password" value={pw1} onChange={(e) => setPw1(e.target.value)} placeholder="Nueva contraseña" className="border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" />
          <input type="password" value={pw2} onChange={(e) => setPw2(e.target.value)} placeholder="Confirmar contraseña" className="border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" />
          <button onClick={changePassword} className="bg-stone-800 hover:bg-stone-900 text-white text-sm font-medium rounded-lg px-4 py-2">Actualizar contraseña</button>
        </div>
      </div>

      {msg && <p className="text-sm text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2">{msg}</p>}
    </div>
  );
}
