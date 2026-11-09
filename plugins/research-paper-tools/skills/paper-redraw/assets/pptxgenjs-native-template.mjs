// Native schematic template, adapted from SciToolsmith; see ../LICENSE.
// Modified for paper-redraw: uniform scaling, unified nodes, single-path edges,
// endpoint validation and optional numeric JSON maps. No image input is read.
// Node.js 20+, PptxGenJS 4.0.x.
// node build.mjs [--map figure.json] [--output editable.pptx] [--overwrite]
// Relative paths resolve from this script. Edges are editable geometry, NOT
// node-bound connectors; use a host-supported attached API if binding is needed.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import PptxGenJS from "pptxgenjs";

const BUILD_DIR = path.dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
if (args.includes("--help")) {
  console.log("node build.mjs [--map figure.json] [--output editable.pptx] [--overwrite]\nMap: source:{width,height}, nodes:[{id,type,text,box:{x,y,w,h}}], edges:[{id,from,to,direction,points:[{x,y}]}]. Optional: slide, fontFace, language, anchorTolerance. Paths resolve from the script; edges are not node-bound.");
  process.exit(0);
}
function option(name) {
  const index = args.indexOf(name);
  if (index < 0) return undefined;
  if (!args[index + 1] || args[index + 1].startsWith("--")) throw new Error(`${name} needs a value`);
  return args[index + 1];
}
const local = (value) => path.resolve(BUILD_DIR, value);
const mapPath = option("--map");
const map = mapPath ? JSON.parse(fs.readFileSync(local(mapPath), "utf8")) : {};
const OUTPUT_PATH = local(option("--output") ?? "editable.pptx");
const SOURCE = map.source ?? { width: 1200, height: 675 };
const SLIDE = {
  width: map.slide?.width ?? 13.333,
  height: map.slide?.height ?? (map.slide?.width ?? 13.333) * SOURCE.height / SOURCE.width,
};
const FONT_FACE = map.fontFace ?? "Arial";
const LANGUAGE = map.language ?? "en-US";
const NODES = map.nodes ?? [
  { id: "question", type: "roundRect", text: "Question",
    box: { x: 90, y: 235, w: 235, h: 92 }, fill: "E8F1FF", line: "3B6CC0" },
  { id: "mechanism", type: "rect", text: "Mechanism",
    box: { x: 485, y: 120, w: 250, h: 92 }, fill: "E7F7F6", line: "238C8C" },
  { id: "outcome", type: "diamond", text: "Outcome",
    box: { x: 485, y: 410, w: 250, h: 120 }, fill: "F4EEFF", line: "7557C8" },
];
const EDGES = map.edges ?? [
  { id: "question-to-mechanism", from: "question", to: "mechanism", direction: "forward",
    points: [{ x: 325, y: 281 }, { x: 405, y: 281 }, { x: 405, y: 166 }, { x: 485, y: 166 }],
    label: "frames", labelBox: { x: 345, y: 212, w: 120, h: 34 } },
  { id: "feedback", from: "mechanism", to: "outcome", direction: "both",
    points: [{ x: 610, y: 212 }, { x: 610, y: 410 }],
    label: "tests", labelBox: { x: 625, y: 290, w: 90, h: 34 } },
];
const SHAPES = { rect: "rect", roundRect: "roundRect", ellipse: "ellipse", diamond: "diamond",
  cube: "cube", can: "can", cloud: "cloud", trapezoid: "trapezoid", parallelogram: "parallelogram" };
const SCALE = Math.min(SLIDE.width / SOURCE.width, SLIDE.height / SOURCE.height);
const OFFSET = { x: (SLIDE.width - SOURCE.width * SCALE) / 2,
  y: (SLIDE.height - SOURCE.height * SCALE) / 2 };
const toPoint = (p) => ({ x: OFFSET.x + p.x * SCALE, y: OFFSET.y + p.y * SCALE });
const toBox = (b) => ({ ...toPoint(b), w: b.w * SCALE, h: b.h * SCALE });
const finite = (n) => typeof n === "number" && Number.isFinite(n);
function fail(message) { throw new Error(`paper-redraw: ${message}`); }

function validateBox(b, label) {
  if (!b || ![b.x, b.y, b.w, b.h].every(finite) || b.w <= 0 || b.h <= 0) fail(`${label}: invalid box`);
  if (b.x < 0 || b.y < 0 || b.x + b.w > SOURCE.width || b.y + b.h > SOURCE.height) fail(`${label}: box exceeds canvas`);
}
function nearBoundary(p, node, tolerance) {
  const b = node.box, dx = p.x - b.x - b.w / 2, dy = p.y - b.y - b.h / 2;
  const rx = b.w / 2, ry = b.h / 2;
  if (node.type === "ellipse") return Math.abs(Math.hypot(dx / rx, dy / ry) - 1) * Math.min(rx, ry) <= tolerance;
  if (node.type === "diamond") return Math.abs(Math.abs(dx) / rx + Math.abs(dy) / ry - 1) / Math.hypot(1 / rx, 1 / ry) <= tolerance;
  if (p.x < b.x - tolerance || p.x > b.x + b.w + tolerance
    || p.y < b.y - tolerance || p.y > b.y + b.h + tolerance) return false;
  return Math.min(Math.abs(p.x - b.x), Math.abs(p.x - b.x - b.w),
    Math.abs(p.y - b.y), Math.abs(p.y - b.y - b.h)) <= tolerance;
}
function validateMap() {
  if (![SOURCE.width, SOURCE.height, SLIDE.width, SLIDE.height].every((n) => finite(n) && n > 0)) fail("invalid canvas");
  const tolerance = map.anchorTolerance ?? Math.max(4, Math.min(SOURCE.width, SOURCE.height) * 0.01);
  if (!finite(tolerance) || tolerance < 0) fail("invalid anchorTolerance");
  const nodes = new Map();
  for (const node of NODES) {
    if (!node.id || nodes.has(node.id)) fail(`duplicate or empty node: ${node.id}`);
    if (!SHAPES[node.type]) fail(`unsupported node type: ${node.type}`);
    validateBox(node.box, node.id);
    nodes.set(node.id, node);
  }
  const ids = new Set();
  for (const edge of EDGES) {
    if (!edge.id || ids.has(edge.id)) fail(`duplicate or empty edge: ${edge.id}`);
    ids.add(edge.id);
    if (!nodes.has(edge.from) || !nodes.has(edge.to)) fail(`${edge.id}: unknown node`);
    if (!["forward", "both", "none"].includes(edge.direction)) fail(`${edge.id}: invalid direction`);
    if (!Array.isArray(edge.points) || edge.points.length < 2) fail(`${edge.id}: explicit path required`);
    for (const p of edge.points) {
      if (!p || !finite(p.x) || !finite(p.y) || p.x < 0 || p.y < 0
        || p.x > SOURCE.width || p.y > SOURCE.height) fail(`${edge.id}: invalid path point`);
    }
    if (!nearBoundary(edge.points[0], nodes.get(edge.from), tolerance)
      || !nearBoundary(edge.points.at(-1), nodes.get(edge.to), tolerance)) fail(`${edge.id}: endpoints do not match from/to node boundaries`);
    if (edge.points.every((p) => p.x === edge.points[0].x && p.y === edge.points[0].y)) fail(`${edge.id}: zero-length path`);
    if (edge.label) validateBox(edge.labelBox, `${edge.id} label`);
  }
}
function addEdge(slide, edge) {
  const points = edge.points.map(toPoint);
  const xs = points.map((p) => p.x), ys = points.map((p) => p.y);
  const x = Math.min(...xs), y = Math.min(...ys);
  const w = Math.max(...xs) - x, h = Math.max(...ys) - y;
  const line = { color: edge.color ?? "3E4652", width: edge.width ?? 1.5,
    dashType: edge.dashType ?? "solid", beginArrowType: edge.direction === "both" ? "triangle" : "none",
    endArrowType: edge.direction === "none" ? "none" : "triangle" };
  if (points.length === 2) {
    slide.addShape("line", { x, y, w, h, objectName: `Edge:${edge.id}`,
      flipH: points[1].x < points[0].x, flipV: points[1].y < points[0].y, line });
  } else {
    slide.addShape("custGeom", { x, y, w: Math.max(w, 0.001), h: Math.max(h, 0.001),
      points: points.map((p) => ({ x: p.x - x, y: p.y - y })), objectName: `Edge:${edge.id}`,
      line, fill: { color: "FFFFFF", transparency: 100 } });
  }
  if (edge.label) slide.addText(edge.label, { ...toBox(edge.labelBox),
    objectName: `Edge-label:${edge.id}`, fontFace: FONT_FACE, fontSize: edge.fontSize ?? 10,
    color: "3E4652", align: "center", valign: "mid", margin: 1.5,
    fill: { color: "FFFFFF", transparency: 8 }, line: { transparency: 100 } });
}
function addNode(slide, node) {
  slide.addText(node.text ?? "", { ...toBox(node.box), shape: SHAPES[node.type], objectName: `Node:${node.id}`,
    fill: { color: node.fill ?? "FFFFFF" }, line: { color: node.line ?? "5B6472", width: node.lineWidth ?? 1.25 },
    fontFace: node.fontFace ?? FONT_FACE, fontSize: node.fontSize ?? 17, bold: node.bold ?? true,
    color: node.color ?? "172033", margin: node.margin ?? 4,
    align: node.align ?? "center", valign: node.valign ?? "mid", lang: LANGUAGE });
}
async function main() {
  if (fs.existsSync(OUTPUT_PATH) && !args.includes("--overwrite")) fail("output exists; use --overwrite for intended replacement");
  validateMap();
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SOURCE", width: SLIDE.width, height: SLIDE.height });
  pptx.layout = "SOURCE";
  pptx.author = "paper-redraw";
  pptx.subject = "Native editable scientific schematic";
  pptx.title = "Paper figure rebuild";
  pptx.lang = LANGUAGE;
  pptx.theme = { headFontFace: FONT_FACE, bodyFontFace: FONT_FACE, lang: LANGUAGE };
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };
  for (const edge of EDGES) addEdge(slide, edge);
  for (const node of NODES) addNode(slide, node);
  await pptx.writeFile({ fileName: OUTPUT_PATH, compression: true });
  process.stdout.write(JSON.stringify({ output: OUTPUT_PATH, slides: 1, runtime: "pptxgenjs", nodeBound: false }) + "\n");
}
main().catch((error) => { process.stderr.write(error.message + "\n"); process.exitCode = 1; });
