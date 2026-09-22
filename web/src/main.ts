import * as XLSX from "xlsx";

type Policy = { upTo: number; reorderPoint: number; reviewPeriod: number };
type Result = { total: number; holding: number; shortage: number; ordering: number };
type Params = { days: number };
type UniformCursor = () => number;

let excelUniforms: number[] | null = null;
let loadedFileName = "";
let excelSource = "";

const money = new Intl.NumberFormat("es-CO", { style: "currency", currency: "USD", maximumFractionDigits: 2 });

function excelCursor(values: number[]): UniformCursor {
  let index = 0;
  return () => {
    if (index >= values.length) throw new Error(`La secuencia de Excel se agotó después de ${values.length.toLocaleString("es-CO")} valores.`);
    const value = values[index];
    index += 1;
    return value;
  };
}

function binomial(nextUniform: UniformCursor): number {
  let value = 0;
  for (let i = 0; i < 6; i += 1) if (nextUniform() < 0.5) value += 1;
  return value;
}

function poisson(nextUniform: UniformCursor): number {
  const uniform = nextUniform();
  let probability = Math.exp(-3);
  let cumulative = probability;
  let value = 0;
  while (uniform > cumulative) {
    value += 1;
    probability *= 3 / value;
    cumulative += probability;
  }
  return value;
}

function simulate(params: Params, policy: Policy, nextUniform: UniformCursor): Result {
  let inventory = 30;
  const arrivals = new Map<number, number>();
  const result: Result = { total: 0, holding: 0, shortage: 0, ordering: 0 };

  for (let day = 1; day <= params.days; day += 1) {
    inventory += arrivals.get(day) ?? 0;
    arrivals.delete(day);
    inventory -= binomial(nextUniform);
    if (inventory >= 0) result.holding += inventory;
    else result.shortage += -inventory * 10;

    const tocaRevisar = policy.reviewPeriod > 0 ? day % policy.reviewPeriod === 0 : inventory <= policy.reorderPoint;
    if (tocaRevisar && inventory < policy.upTo) {
      const quantity = policy.upTo - inventory;
      const arrivalDay = day + Math.max(1, poisson(nextUniform));
      arrivals.set(arrivalDay, (arrivals.get(arrivalDay) ?? 0) + quantity);
      result.ordering += 50;
    }
  }
  result.total = result.holding + result.shortage + result.ordering;
  return result;
}

function setText(id: string, value: string): void { document.getElementById(id)!.textContent = value; }
function resetResults(message: string): void {
  for (const id of ["p1-total", "p1-daily", "p1-holding", "p1-shortage", "p1-ordering", "p2-total", "p2-daily", "p2-holding", "p2-shortage", "p2-ordering"]) setText(id, "—");
  const verdict = document.getElementById("verdict")!;
  verdict.classList.add("neutral");
  verdict.textContent = message;
}
function showResult(prefix: "p1" | "p2", result: Result, days: number): void {
  setText(`${prefix}-total`, money.format(result.total));
  setText(`${prefix}-daily`, `${money.format(result.total / days)} por día`);
  setText(`${prefix}-holding`, money.format(result.holding));
  setText(`${prefix}-shortage`, money.format(result.shortage));
  setText(`${prefix}-ordering`, money.format(result.ordering));
}

function run(): void {
  const params: Params = { days: Number((document.getElementById("days") as HTMLInputElement).value) };
  if (!Number.isInteger(params.days) || params.days < 1) { setText("verdict", "Ingresa un número entero de días mayor que cero."); return; }
  if (!excelUniforms) { resetResults("Carga un archivo Excel válido antes de ejecutar."); return; }
  if (params.days * 7 > excelUniforms.length) { resetResults(`El archivo aporta ${excelUniforms.length.toLocaleString("es-CO")} R_i. Para ${params.days} días se requieren hasta ${(params.days * 7).toLocaleString("es-CO")}.`); return; }
  try {
    const first = simulate(params, { upTo: 30, reorderPoint: 0, reviewPeriod: 8 }, excelCursor(excelUniforms));
    const second = simulate(params, { upTo: 30, reorderPoint: 10, reviewPeriod: 0 }, excelCursor(excelUniforms));
    showResult("p1", first, params.days); showResult("p2", second, params.days);
    const [winner, loser, saving] = first.total <= second.total
      ? ["Política 1", "Política 2", second.total - first.total]
      : ["Política 2", "Política 1", first.total - second.total];
    document.getElementById("verdict")!.classList.remove("neutral");
    setText("verdict", `${winner} es más económica: ahorra ${money.format(saving)} frente a ${loser}.`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "No fue posible consumir los números del archivo.";
    setText("verdict", `${message} Reduce los días o sube una secuencia más larga.`);
  }
}

function normaliseHeader(value: unknown): string {
  return String(value ?? "").toLowerCase().replace(/[\s_\-()]/g, "");
}

function uniformsFromGeneratorParameters(rows: unknown[][]): { values: number[]; source: string } | null {
  const parameters = new Map<string, number>();
  for (const row of rows) {
    const label = normaliseHeader(row[5]);
    const value = Number(row[6]);
    if (label && Number.isFinite(value)) parameters.set(label, value);
  }
  const seed = parameters.get("semillax0");
  const a = parameters.get("a");
  const c = parameters.get("c");
  const m = parameters.get("m");
  const count = parameters.get("n");
  if ([seed, a, m, count].some((value) => value === undefined) || !Number.isInteger(m) || !Number.isInteger(count) || m! <= 1 || count! < 1) return null;
  const values: number[] = [];
  let x = seed!;
  for (let i = 0; i < count!; i += 1) {
    x = (a! * x + (c ?? 0)) % m!;
    values.push(x / m!);
  }
  return { values, source: c === undefined ? "parámetros lineales multiplicativos" : "parámetros lineales mixtos" };
}

async function importExcel(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: true });
    const sheetName = workbook.SheetNames.find((name) => normaliseHeader(name) === "linealmixto") ?? workbook.SheetNames[0];
    if (!sheetName) throw new Error("El libro no contiene hojas.");
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true });
    const headerRow = rows.findIndex((row) => row.some((cell) => normaliseHeader(cell) === "ri"));
    if (headerRow < 0) throw new Error("No se encontró la columna R_i.");
    const column = rows[headerRow].findIndex((cell) => normaliseHeader(cell) === "ri");
    const cells = rows.slice(headerRow + 1).map((row) => row[column]).filter((value) => value !== null && value !== "");
    let values = cells.map(Number);
    excelSource = "valores de la columna R_i";
    if (!values.length || values.some((value) => !Number.isFinite(value) || value < 0 || value >= 1)) {
      const generated = uniformsFromGeneratorParameters(rows);
      if (!generated) throw new Error("R_i debe traer números entre 0 y 1, o el archivo debe incluir parámetros lineales válidos (X0, a, m, N y opcionalmente c).");
      values = generated.values;
      excelSource = `${generated.source} del archivo (fórmulas reconstruidas)`;
    }
    if (values.some((value) => value < 0 || value >= 1)) throw new Error("R_i debe contener números entre 0 (incluido) y 1 (excluido).");
    excelUniforms = values;
    loadedFileName = file.name;
    const days = document.getElementById("days") as HTMLInputElement;
    const safeDays = Math.max(1, Math.floor(values.length / 7));
    days.value = String(safeDays);
    setText("excel-status", `${loadedFileName}: ${values.length.toLocaleString("es-CO")} R_i obtenidos de ${excelSource} en “${sheetName}”. Se establecieron ${days.value} días para no reutilizarlos.`);
    (document.getElementById("run") as HTMLButtonElement).disabled = false;
    resetResults("Archivo cargado. El sistema usará únicamente estos R_i; pulsa “Ejecutar simulación”.");
    setWalkthroughFromUniforms(values);
  } catch (error) {
    excelUniforms = null;
    loadedFileName = "";
    excelSource = "";
    (document.getElementById("run") as HTMLButtonElement).disabled = true;
    const message = error instanceof Error ? error.message : "No se pudo leer el archivo.";
    setText("excel-status", message);
    resetResults("El archivo no se cargó. Revisa que tenga la columna R_i.");
  }
}

type WalkStep = { label: string; title: string; before: string; operator: string; after: string; explanation: string; summary: string; used: number; code: number };
const sampleWalkthrough: WalkStep[] = [
  { label: "Inicio", title: "Inventario inicial", before: "30", operator: "unidades disponibles", after: "30", explanation: "Se inicia con 30 unidades. No hay pedidos que recibir todavía.", summary: "Inventario neto: 30", used: 0, code: 0 },
  { label: "Paso 1 · recibir", title: "Llegada de pedidos", before: "30 + 0", operator: "pedido recibido", after: "30", explanation: "Se suman primero las entregas programadas para este día. En el día 1 no llega ninguna.", summary: "Inventario neto: 30", used: 0, code: 1 },
  { label: "Paso 2 · demandar", title: "Demanda Binomial(6, 0,5)", before: "30 − 3", operator: "tres R_i < 0,5", after: "27", explanation: "Se leen R₁ a R₆. Los valores 0,12, 0,46 y 0,31 son menores que 0,5, así que la demanda es 3.", summary: "Inventario neto: 27", used: 6, code: 2 },
  { label: "Paso 3 · costear", title: "Costo de mantenimiento", before: "27 × $1", operator: "inventario final", after: "$27", explanation: "Como el inventario final es positivo, se cobra $1 por unidad. No hay costo de faltante este día.", summary: "Costo acumulado del día: $27", used: 6, code: 3 },
  { label: "Paso 4 · ordenar", title: "Política 1: revisar cada 8 días", before: "día 1", operator: "1 % 8 ≠ 0", after: "no se revisa", explanation: "La política 1 solo revisa el inventario cada 8 días. Como el día 1 no es múltiplo de 8, no se evalúa si hay que ordenar, sin importar el nivel de inventario.", summary: "Costo del día: $27 (sin orden)", used: 6, code: 4 },
];

let walkthrough = sampleWalkthrough;
let walkValues = ["R₁ 0,12", "R₂ 0,88", "R₃ 0,46", "R₄ 0,67", "R₅ 0,31", "R₆ 0,75"];
let walkCurrent = 0;
let walkUniforms: HTMLElement | null = null;

function formatUniform(value: number): string { return value.toLocaleString("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 3 }); }

function renderWalkthrough(): void {
  if (!walkUniforms) return;
  const step = walkthrough[walkCurrent];
  walkUniforms.innerHTML = walkValues.map((value, index) => `<span class="${index < step.used ? "consumed" : ""}">${value}</span>`).join("");
  setText("walk-count", `Paso ${walkCurrent + 1} de ${walkthrough.length}`); setText("walk-label", step.label); setText("walk-title", step.title);
  setText("walk-before", step.before); setText("walk-operator", step.operator); setText("walk-after", step.after); setText("walk-explanation", step.explanation); setText("walk-summary", step.summary);
  document.querySelectorAll<HTMLLIElement>("#walk-code li").forEach((element, index) => element.classList.toggle("active", index === step.code - 1));
  (document.getElementById("walk-back") as HTMLButtonElement).disabled = walkCurrent === 0;
  (document.getElementById("walk-next") as HTMLButtonElement).disabled = walkCurrent === walkthrough.length - 1;
}

function setWalkthroughFromUniforms(values: number[]): void {
  const demandUniforms = values.slice(0, 6);
  const demand = demandUniforms.filter((value) => value < 0.5).length;
  const inventory = 30 - demand;
  const labels = demandUniforms.map((value, index) => `R${index + 1} ${formatUniform(value)}`);
  const marked = demandUniforms.filter((value) => value < 0.5).map(formatUniform).join(", ") || "ninguno";
  const steps: WalkStep[] = [
    { label: "Día 1 · inicio", title: "Inventario inicial", before: "30", operator: "unidades disponibles", after: "30", explanation: "La corrida comienza con 30 unidades, como está definido en el modelo.", summary: "Inventario neto: 30", used: 0, code: 0 },
    { label: "Paso 1 · recibir", title: "Llegada de pedidos", before: "30 + 0", operator: "pedido recibido", after: "30", explanation: "En el primer día no existe una orden anterior que pueda llegar.", summary: "Inventario neto: 30", used: 0, code: 1 },
    { label: "Paso 2 · demandar", title: "Demanda del archivo", before: `30 − ${demand}`, operator: `${demand} R_i < 0,5`, after: String(inventory), explanation: `Se consumen R₁–R₆ del Excel. Los menores que 0,5 son: ${marked}; por eso la demanda es ${demand}.`, summary: `Inventario neto: ${inventory}`, used: 6, code: 2 },
    { label: "Paso 3 · costear", title: "Costo de mantenimiento", before: `${inventory} × $1`, operator: "inventario final", after: `$${inventory}`, explanation: `El inventario es positivo, por lo que se cobra $1 por cada una de las ${inventory} unidades.`, summary: `Costo acumulado del día: $${inventory}`, used: 6, code: 3 },
    { label: "Paso 4 · ordenar", title: "Política 1: revisar cada 8 días", before: "día 1", operator: "1 % 8 ≠ 0", after: "no se revisa", explanation: "La política 1 solo evalúa si debe ordenar cada 8 días. Como el día 1 no es múltiplo de 8, no se revisa el inventario, sin importar cuánto quedó tras la demanda.", summary: `Costo del día: $${inventory} (sin orden)`, used: 6, code: 4 },
  ];
  walkthrough = steps; walkValues = labels; walkCurrent = 0;
  setText("guided-note", "Este recorrido usa los primeros R_i del archivo cargado y representa la política 1 durante el día 1.");
  renderWalkthrough();
}

function initialiseWalkthrough(): void {
  walkUniforms = document.getElementById("walk-uniforms");
  document.getElementById("walk-back")!.addEventListener("click", () => { walkCurrent -= 1; renderWalkthrough(); });
  document.getElementById("walk-next")!.addEventListener("click", () => { walkCurrent += 1; renderWalkthrough(); });
  renderWalkthrough();
}

document.getElementById("run")!.addEventListener("click", run);
document.getElementById("excel")!.addEventListener("change", importExcel);
initialiseWalkthrough();
