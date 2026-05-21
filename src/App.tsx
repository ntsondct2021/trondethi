import React, { useState, useEffect, useRef, useMemo } from 'react';
import { 
  LayoutDashboard, 
  Database, 
  Settings, 
  FileText, 
  Shuffle, 
  Download, 
  Plus, 
  Search, 
  Filter, 
  Trash2, 
  Edit, 
  Upload,
  ChevronRight,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Sparkles,
  Wand2,
  Image as ImageIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import JSZip from 'jszip';
import { Question, QuestionType, ExamStructure, GeneratedExam } from './types';
import { mixExams } from './randomizer';
import { saveAs } from 'file-saver';
import { generateQuestionsWithAI } from './services/gemini';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  AlignmentType, 
  HeadingLevel, 
  ImageRun, 
  Header, 
  Footer, 
  PageNumber,
  Math as DocxMath,
  MathRun,
  MathFraction,
  MathSuperScript,
  MathSubScript,
  MathSubSuperScript,
  MathRadical,
  MathRoundBrackets,
  MathCurlyBrackets,
  MathSquareBrackets,
  MathComponent
} from 'docx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import ReactQuill from 'react-quill-new';
import 'react-quill-new/dist/quill.snow.css';

// --- OMML & MathType to LaTeX Helpers for Word Import ---

interface Token {
  type: 'LBRACE' | 'RBRACE' | 'LBRACKET' | 'RBRACKET' | 'LPAREN' | 'RPAREN' | 'SUB' | 'SUP' | 'CMD' | 'TEXT';
  value: string;
}

const findChildByLocalName = (node: Node, localName: string): Element | null => {
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (child.nodeType === Node.ELEMENT_NODE && (child as Element).localName === localName) {
      return child as Element;
    }
  }
  return null;
};

const convertOmmlNodeToLatex = (node: Node): string => {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || '';
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return '';
  }

  const el = node as Element;
  const name = el.localName;

  const getChildContent = (childName: string): string => {
    const child = findChildByLocalName(el, childName);
    return child ? convertOmmlNodeToLatex(child) : '';
  };

  const getChildrenContent = (element: Element): string => {
    let res = '';
    for (let i = 0; i < element.childNodes.length; i++) {
      res += convertOmmlNodeToLatex(element.childNodes[i]);
    }
    return res;
  };

  switch (name) {
    case 't':
      return el.textContent || '';
    case 'r':
    case 'oMath':
    case 'oMathPara':
    case 'e':
    case 'lim':
    case 'num':
    case 'den':
    case 'sub':
    case 'sup':
    case 'deg':
      return getChildrenContent(el);

    case 'f': {
      const numStr = getChildContent('num');
      const denStr = getChildContent('den');
      return `\\frac{${numStr}}{${denStr}}`;
    }

    case 'sSup': {
      const baseStr = getChildContent('e');
      const supStr = getChildContent('sup');
      return `{${baseStr}}^{${supStr}}`;
    }

    case 'sSub': {
      const baseStr = getChildContent('e');
      const subStr = getChildContent('sub');
      return `{${baseStr}}_{${subStr}}`;
    }

    case 'sSubSup': {
      const baseStr = getChildContent('e');
      const subStr = getChildContent('sub');
      const supStr = getChildContent('sup');
      return `{${baseStr}}_{${subStr}}^{${supStr}}`;
    }

    case 'rad': {
      const baseStr = getChildContent('e');
      const degStr = getChildContent('deg');
      if (degStr.trim()) {
        return `\\sqrt[${degStr}]{${baseStr}}`;
      }
      return `\\sqrt{${baseStr}}`;
    }

    case 'd': {
      const eStr = getChildContent('e');
      let begChr = '(';
      let endChr = ')';
      
      const dPr = findChildByLocalName(el, 'dPr');
      if (dPr) {
        const begChrEl = findChildByLocalName(dPr, 'begChr');
        const endChrEl = findChildByLocalName(dPr, 'endChr');
        if (begChrEl) begChr = begChrEl.getAttribute('val') || '(';
        if (endChrEl) endChr = endChrEl.getAttribute('val') || ')';
      }

      if (begChr === '{') begChr = '\\{';
      if (endChr === '}') endChr = '\\}';

      return `\\left${begChr}${eStr}\\right${endChr}`;
    }

    case 'nary': {
      const eStr = getChildContent('e');
      const subStr = getChildContent('sub');
      const supStr = getChildContent('sup');
      
      let op = '\\sum';
      const naryPr = findChildByLocalName(el, 'naryPr');
      if (naryPr) {
        const chrEl = findChildByLocalName(naryPr, 'chr');
        if (chrEl) {
          const charVal = chrEl.getAttribute('val') || '∑';
          if (charVal === '∫' || charVal === 'int') op = '\\int';
          else if (charVal === '∏' || charVal === 'prod') op = '\\prod';
        }
      }

      let limits = '';
      if (subStr.trim()) limits += `_{${subStr}}`;
      if (supStr.trim()) limits += `^{${supStr}}`;
      
      return `${op}${limits}{${eStr}}`;
    }

    case 'limLow': {
      const baseStr = getChildContent('e');
      const limStr = getChildContent('lim');
      return `{${baseStr}}_{${limStr}}`;
    }

    case 'limUpp': {
      const baseStr = getChildContent('e');
      const limStr = getChildContent('lim');
      return `{${baseStr}}^{${limStr}}`;
    }

    case 'm': {
      let rowsStr = '';
      const rows: Element[] = [];
      for (let i = 0; i < el.childNodes.length; i++) {
        const child = el.childNodes[i];
        if (child.nodeType === Node.ELEMENT_NODE && (child as Element).localName === 'mr') {
          rows.push(child as Element);
        }
      }
      
      const rowLatex = rows.map(row => {
        const cells: Element[] = [];
        for (let j = 0; j < row.childNodes.length; j++) {
          const child2 = row.childNodes[j];
          if (child2.nodeType === Node.ELEMENT_NODE && (child2 as Element).localName === 'e') {
            cells.push(child2 as Element);
          }
        }
        return cells.map(cell => convertOmmlNodeToLatex(cell)).join(' & ');
      }).join(' \\\\ ');

      return `\\begin{matrix}${rowLatex}\\end{matrix}`;
    }

    default:
      return getChildrenContent(el);
  }
};

const arrayBufferToBase64 = (buffer: ArrayBuffer | Uint8Array): string => {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
};

const base64ToUint8Array = (base64: string): Uint8Array => {
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
};

const preprocessDocx = async (arrayBuffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentXmlText = await zip.file("word/document.xml")?.async("string");
  if (!documentXmlText) return arrayBuffer;

  const parser = new DOMParser();
  const doc = parser.parseFromString(documentXmlText, "application/xml");

  const relsXmlText = await zip.file("word/_rels/document.xml.rels")?.async("string");
  const relsMap = new Map<string, string>();
  if (relsXmlText) {
    const relsDoc = parser.parseFromString(relsXmlText, "application/xml");
    const rels = relsDoc.getElementsByTagName("Relationship");
    for (let i = 0; i < rels.length; i++) {
      const rel = rels[i];
      const id = rel.getAttribute("Id");
      const target = rel.getAttribute("Target");
      if (id && target) {
        relsMap.set(id, target.startsWith("word/") ? target : `word/${target}`);
      }
    }
  }

  const mathAssets: any[] = [];
  let mathCounter = 1;
  const batchId = Date.now() + "_" + Math.floor(Math.random() * 1000);

  let modified = false;

  // 1. Process standard oMath elements (Office Math)
  const oMathNodes: Element[] = [];
  const allElements = doc.getElementsByTagName("*");
  for (let i = 0; i < allElements.length; i++) {
    const el = allElements[i];
    if (el.localName === 'oMath') {
      oMathNodes.push(el);
    }
  }

  for (let i = oMathNodes.length - 1; i >= 0; i--) {
    const oMath = oMathNodes[i];
    let parent = oMath.parentNode;
    let isNested = false;
    while (parent) {
      if ((parent as Element).localName === 'oMath') {
        isNested = true;
        break;
      }
      parent = parent.parentNode;
    }
    if (isNested) continue;

    // Check if oMath is wrapped in an OLE object (MathType). If so, let OLE object parsing handle it.
    let isWrappedInObject = false;
    let p = oMath.parentNode;
    while (p) {
      if ((p as Element).localName === 'object' || (p as Element).localName === 'objectPr') {
        isWrappedInObject = true;
        break;
      }
      p = p.parentNode;
    }
    if (isWrappedInObject) continue;

    const latex = convertOmmlNodeToLatex(oMath);
    const originalXml = new XMLSerializer().serializeToString(oMath);
    const id = `mathtype_${batchId}_om_${mathCounter++}`;

    mathAssets.push({
      id,
      originalXml,
      latexFallback: latex || ""
    });

    const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const rElement = doc.createElementNS(wNamespace, "w:r");
    const tElement = doc.createElementNS(wNamespace, "w:t");
    tElement.textContent = `[!m:${id}]`;
    rElement.appendChild(tElement);

    oMath.parentNode?.replaceChild(rElement, oMath);
    modified = true;
  }

  // 2. Process OLE Objects and Shapes (MathType elements)
  const objectNodes: Element[] = [];
  for (let i = 0; i < allElements.length; i++) {
    const el = allElements[i];
    if (el.localName === 'object') {
      objectNodes.push(el);
    }
  }

  for (let i = objectNodes.length - 1; i >= 0; i--) {
    const node = objectNodes[i];
    
    // Seek OLE details
    const oleObj = node.getElementsByTagName("o:OLEObject")[0] || node.querySelector("OLEObject");
    if (!oleObj) continue;

    const oleRelId = oleObj.getAttribute("r:id") || oleObj.getAttribute("id");
    if (!oleRelId) continue;

    const oleTargetPath = relsMap.get(oleRelId);
    if (!oleTargetPath) continue;

    // Seek Image Details
    const imagedata = node.getElementsByTagName("v:imagedata")[0] || node.querySelector("imagedata");
    let imgRelId = imagedata ? (imagedata.getAttribute("r:id") || imagedata.getAttribute("id")) : null;
    let imgTargetPath = imgRelId ? relsMap.get(imgRelId) : null;

    let oleBinBase64 = "";
    const oleFile = zip.file(oleTargetPath);
    if (oleFile) {
      const bytes = await oleFile.async("uint8array");
      oleBinBase64 = arrayBufferToBase64(bytes);
    }

    let imageBinBase64 = "";
    let imageExt = "png";
    if (imgTargetPath) {
      const imgFile = zip.file(imgTargetPath);
      if (imgFile) {
        const bytes = await imgFile.async("uint8array");
        imageBinBase64 = arrayBufferToBase64(bytes);
        imageExt = imgTargetPath.split('.').pop() || "png";
      }
    }

    // Extract LaTeX fallback from shape alt/descr/title
    let latexFallback = "";
    const alt = node.getAttribute('alt') || node.getAttribute('descr') || node.getAttribute('title');
    const shape = node.getElementsByTagName("v:shape")[0] || node.querySelector("shape");
    const shapeAlt = shape ? (shape.getAttribute('alt') || shape.getAttribute('descr') || shape.getAttribute('o:title') || shape.getAttribute('title')) : '';
    const desc = alt || shapeAlt || '';
    if (desc.includes('MathType LaTeX Equation:')) {
      latexFallback = desc.split('MathType LaTeX Equation:')[1].trim();
    } else if (desc.includes('Equation:')) {
      latexFallback = desc.split('Equation:')[1].trim();
    } else if (desc.startsWith('$') && desc.endsWith('$')) {
      latexFallback = desc.slice(1, -1).trim();
    }

    const originalXml = new XMLSerializer().serializeToString(node);
    const id = `mathtype_${batchId}_ole_${mathCounter++}`;

    mathAssets.push({
      id,
      originalXml,
      oleBinBase64,
      imageBinBase64,
      imageExt,
      latexFallback
    });

    // Find the outermore run context target to replace
    let targetToReplace: Node = node;
    let parent = node.parentNode;
    while (parent) {
      const name = (parent as Element).localName || parent.nodeName;
      if (name === 'r' || name === 'drawing' || name === 'object') {
        targetToReplace = parent;
      }
      parent = parent.parentNode;
    }

    const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const rElement = doc.createElementNS(wNamespace, "w:r");
    const tElement = doc.createElementNS(wNamespace, "w:t");
    tElement.textContent = `[!m:${id}]`;
    rElement.appendChild(tElement);

    targetToReplace.parentNode?.replaceChild(rElement, targetToReplace);
    modified = true;
  }

  // 3. Process drawings (standard drawings/images)
  const drawingNodes: Element[] = [];
  const currentElements3 = doc.getElementsByTagName("*");
  for (let i = 0; i < currentElements3.length; i++) {
    const el = currentElements3[i];
    if (el.localName === 'drawing') {
      drawingNodes.push(el);
    }
  }

  for (let i = drawingNodes.length - 1; i >= 0; i--) {
    const drawing = drawingNodes[i];
    if (!doc.contains(drawing)) continue;

    let isWrappedInObject = false;
    let p = drawing.parentNode;
    while (p) {
      const name = (p as Element).localName || p.nodeName;
      if (name === 'object' || name === 'objectPr') {
        isWrappedInObject = true;
        break;
      }
      p = p.parentNode;
    }
    if (isWrappedInObject) continue;

    const blips = drawing.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "blip");
    let embedId = null;
    if (blips && blips.length > 0) {
      embedId = blips[0].getAttribute("r:embed") || blips[0].getAttribute("embed") || blips[0].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
    } else {
      const anyBlips = drawing.getElementsByTagName("blip") || drawing.querySelectorAll("blip");
      if (anyBlips && anyBlips.length > 0) {
        embedId = anyBlips[0].getAttribute("r:embed") || anyBlips[0].getAttribute("embed");
      }
    }

    if (!embedId) continue;

    const imgTargetPath = relsMap.get(embedId);
    if (!imgTargetPath) continue;

    const imgFile = zip.file(imgTargetPath);
    if (!imgFile) continue;

    const bytes = await imgFile.async("uint8array");
    const imageBinBase64 = arrayBufferToBase64(bytes);
    const imageExt = imgTargetPath.split('.').pop()?.toLowerCase() || "png";
    const mimeType = imageExt === "jpg" || imageExt === "jpeg" ? "image/jpeg" : `image/${imageExt}`;

    const dataUrl = `data:${mimeType};base64,${imageBinBase64}`;

    let targetToReplace: Node = drawing;
    let parentNode = drawing.parentNode;
    while (parentNode) {
      const name = (parentNode as Element).localName || parentNode.nodeName;
      if (name === 'r' || name === 'drawing') {
        targetToReplace = parentNode;
      }
      parentNode = parentNode.parentNode;
    }

    const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const rElement = doc.createElementNS(wNamespace, "w:r");
    const tElement = doc.createElementNS(wNamespace, "w:t");
    tElement.textContent = `[!img:${dataUrl}]`;
    rElement.appendChild(tElement);

    targetToReplace.parentNode?.replaceChild(rElement, targetToReplace);
    modified = true;
  }

  // 4. Process legacy VML / standalone imagedata elements
  const imagedataNodes: Element[] = [];
  const currentElements4 = doc.getElementsByTagName("*");
  for (let i = 0; i < currentElements4.length; i++) {
    const el = currentElements4[i];
    if (el.localName === 'imagedata') {
      imagedataNodes.push(el);
    }
  }

  for (let i = imagedataNodes.length - 1; i >= 0; i--) {
    const imagedata = imagedataNodes[i];
    if (!doc.contains(imagedata)) continue;

    let isWrapped = false;
    let p = imagedata.parentNode;
    while (p) {
      const name = (p as Element).localName || p.nodeName;
      if (name === 'object' || name === 'objectPr' || name === 'drawing') {
        isWrapped = true;
        break;
      }
      p = p.parentNode;
    }
    if (isWrapped) continue;

    let imgRelId = imagedata.getAttribute("r:id") || imagedata.getAttribute("id") || imagedata.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    if (!imgRelId) continue;

    const imgTargetPath = relsMap.get(imgRelId);
    if (!imgTargetPath) continue;

    const imgFile = zip.file(imgTargetPath);
    if (!imgFile) continue;

    const bytes = await imgFile.async("uint8array");
    const imageBinBase64 = arrayBufferToBase64(bytes);
    const imageExt = imgTargetPath.split('.').pop()?.toLowerCase() || "png";
    const mimeType = imageExt === "jpg" || imageExt === "jpeg" ? "image/jpeg" : `image/${imageExt}`;

    const dataUrl = `data:${mimeType};base64,${imageBinBase64}`;

    let targetToReplace: Node = imagedata;
    let parentNode = imagedata.parentNode;
    while (parentNode) {
      const name = (parentNode as Element).localName || parentNode.nodeName;
      if (name === 'r' || name === 'shape' || name === 'pict') {
        targetToReplace = parentNode;
      }
      parentNode = parentNode.parentNode;
    }

    const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const rElement = doc.createElementNS(wNamespace, "w:r");
    const tElement = doc.createElementNS(wNamespace, "w:t");
    tElement.textContent = `[!img:${dataUrl}]`;
    rElement.appendChild(tElement);

    targetToReplace.parentNode?.replaceChild(rElement, targetToReplace);
    modified = true;
  }

  // 5. Send all extracted assets to Server
  if (mathAssets.length > 0) {
    try {
      await fetch('/api/math-assets', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(mathAssets)
      });
      console.log(`Extracted and uploaded ${mathAssets.length} math assets to the server.`);
    } catch (saveErr) {
      console.error("Failed to save math assets to server:", saveErr);
    }
  }

  if (modified) {
    const serializer = new XMLSerializer();
    const newXmlText = serializer.serializeToString(doc);
    zip.file("word/document.xml", newXmlText);
    return await zip.generateAsync({ type: "arraybuffer" });
  }

  return arrayBuffer;
};

const postprocessMathAssetsInDocx = async (shuffledBlob: Blob, exam: GeneratedExam): Promise<Blob> => {
  // Load the newly generated docx as zip first to scan for placeholders directly in document.xml
  const zip = await JSZip.loadAsync(shuffledBlob);

  const docXmlText = await zip.file("word/document.xml")?.async("string");
  if (!docXmlText) return shuffledBlob;

  const assetIdsReferenced = new Set<string>();
  const matches = docXmlText.matchAll(/\[!m:(mathtype_[^\]]+?)\]/g);
  for (const match of matches) {
    assetIdsReferenced.add(match[1]);
  }

  const matchesDollar = docXmlText.matchAll(/\[!m:\$(mathtype_[^\$]+?)\$\]/g);
  for (const match of matchesDollar) {
    assetIdsReferenced.add(match[1]);
  }

  console.log("postprocessMathAssetsInDocx discovered asset IDs directly from document.xml:", Array.from(assetIdsReferenced));

  if (assetIdsReferenced.size === 0) {
    return shuffledBlob;
  }

  // Fetch math assets details in bulk
  let fetchedAssets: any[] = [];
  try {
    const response = await fetch('/api/math-assets/get-bulk', {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ ids: Array.from(assetIdsReferenced) })
    });
    fetchedAssets = await response.json();
  } catch (err) {
    console.error("Failed to fetch math assets for post-processing:", err);
    return shuffledBlob;
  }

  if (fetchedAssets.length === 0) {
    return shuffledBlob;
  }

  const assetsMap = new Map<string, any>();
  fetchedAssets.forEach(asset => assetsMap.set(asset.id, asset));

  const parser = new DOMParser();
  const doc = parser.parseFromString(docXmlText, "application/xml");

  // Load and parse relationships
  let relsXmlText = await zip.file("word/_rels/document.xml.rels")?.async("string");
  if (!relsXmlText) {
    relsXmlText = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  }
  const relsDoc = parser.parseFromString(relsXmlText, "application/xml");
  const relationshipsTag = relsDoc.documentElement;

  const rToReplaceList: { rNode: Element, assetId: string }[] = [];

  // Locate all runs containing the placeholder! Case-insensitive + namespace-agnostic matching
  const allDocElements = doc.getElementsByTagName("*");
  for (let i = 0; i < allDocElements.length; i++) {
    const rNode = allDocElements[i];
    const rNodeName = (rNode.localName || rNode.nodeName).toLowerCase();
    if (rNodeName === 'r' || rNodeName === 'w:r') {
      let tNode: Element | null = null;
      for (let j = 0; j < rNode.childNodes.length; j++) {
        const child = rNode.childNodes[j];
        if (child.nodeType === Node.ELEMENT_NODE) {
          const childName = ((child as Element).localName || child.nodeName).toLowerCase();
          if (childName === 't' || childName === 'w:t') {
            tNode = child as Element;
            break;
          }
        }
      }
      if (tNode && tNode.textContent) {
        const contentText = tNode.textContent;
        const match = contentText.match(/\[!m:(mathtype_[^\]]+?)\]/);
        if (match) {
          const assetId = match[1];
          if (assetsMap.has(assetId)) {
            rToReplaceList.push({ rNode, assetId });
          }
        }
      }
    }
  }

  const replacedAssetIds = new Set<string>();

  for (const item of rToReplaceList) {
    const { rNode, assetId } = item;
    const asset = assetsMap.get(assetId);
    if (!asset) continue;

    // Wrap original XML within a comprehensive root of namespace bindings to secure against prefix unbound XML parser failures
    let actualAssetElement: Element | null = null;
    try {
      const wrappedXml = `<wrapper 
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:w10="urn:schemas-microsoft-com:office:word"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
        xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
      >${asset.originalXml}</wrapper>`;

      const assetDoc = parser.parseFromString(wrappedXml, "application/xml");
      const parserError = assetDoc.getElementsByTagName("parsererror")[0];
      if (!parserError) {
        actualAssetElement = assetDoc.documentElement.firstElementChild;
      } else {
        console.warn("XML wrapped Parse Error for asset, trying direct parse:", assetId, parserError.textContent);
        const directDoc = parser.parseFromString(asset.originalXml, "application/xml");
        if (!directDoc.getElementsByTagName("parsererror")[0]) {
          actualAssetElement = directDoc.documentElement;
        }
      }
    } catch (parseErr) {
      console.error("Critical error parsing XML for asset:", assetId, parseErr);
    }

    if (!actualAssetElement) {
      console.error("Failed to parse math asset XML for", assetId);
      continue;
    }

    const assetNode = doc.importNode(actualAssetElement, true);

    // Clean style to make sure the legacy MathType shapes render inline-with-text instead of floating absolutely
    const allAssetElements = assetNode.getElementsByTagName("*");
    for (let j = 0; j < allAssetElements.length; j++) {
      const el = allAssetElements[j];
      const styleAttr = el.getAttribute("style");
      if (styleAttr) {
        const stylePairs = styleAttr.split(';').map(p => p.trim()).filter(p => p !== '');
        const cleanPairs = stylePairs.filter(pair => {
          const parts = pair.split(':');
          if (parts.length < 2) return false;
          const key = parts[0].trim().toLowerCase();
          return ![
            'position',
            'margin-left',
            'margin-top',
            'left',
            'top',
            'mso-position-horizontal',
            'mso-position-vertical',
            'mso-position-horizontal-relative',
            'mso-position-vertical-relative',
            'mso-wrap-style',
            'mso-wrap-distance-left',
            'mso-wrap-distance-top',
            'mso-wrap-distance-right',
            'mso-wrap-distance-bottom'
          ].includes(key);
        });
        cleanPairs.push('position:static');
        el.setAttribute("style", cleanPairs.join(';'));
      }
    }

    // Update relationship references in the assetNode if it's an OLE object
    if (asset.oleBinBase64) {
      // Set relation IDs inside the OLE XML in a namespace-independent manner
      const OLE_REL_ID = `relOle_${assetId}`;
      const IMG_REL_ID = `relImg_${assetId}`;
      const rNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

      for (let j = 0; j < allAssetElements.length; j++) {
        const el = allAssetElements[j];
        const elName = (el.localName || el.nodeName).toLowerCase();
        if (elName === 'oleobject' || elName === 'o:oleobject') {
          el.setAttributeNS(rNamespace, "r:id", OLE_REL_ID);
        }
        if (elName === 'imagedata' || elName === 'v:imagedata') {
          el.setAttributeNS(rNamespace, "r:id", IMG_REL_ID);
        }
      }

      // Add actual relationships to relsDoc if they are not already added using correct Relationships schema namespaces
      if (!replacedAssetIds.has(assetId) && relationshipsTag) {
        const relNs = relationshipsTag.namespaceURI || "http://schemas.openxmlformats.org/package/2006/relationships";

        // Add OLE binary relationship
        const oleRel = relsDoc.createElementNS(relNs, "Relationship");
        oleRel.setAttribute("Id", OLE_REL_ID);
        oleRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject");
        oleRel.setAttribute("Target", `embeddings/ole_${assetId}.bin`);
        relationshipsTag.appendChild(oleRel);

        // Add Image relationship
        const imgRel = relsDoc.createElementNS(relNs, "Relationship");
        imgRel.setAttribute("Id", IMG_REL_ID);
        imgRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        imgRel.setAttribute("Target", `media/img_${assetId}.${asset.imageExt || 'png'}`);
        relationshipsTag.appendChild(imgRel);

        // Write binary assets to JSZip
        const oleBytes = base64ToUint8Array(asset.oleBinBase64);
        zip.file(`word/embeddings/ole_${assetId}.bin`, oleBytes);

        if (asset.imageBinBase64) {
          const imgBytes = base64ToUint8Array(asset.imageBinBase64);
          zip.file(`word/media/img_${assetId}.${asset.imageExt || 'png'}`, imgBytes);
        }

        replacedAssetIds.add(assetId);
      }
    }

    // Wrap w:object in a w:r segment to maintain perfect inline DOCX schema validation and styling
    let finalNodeToInsert = assetNode;
    const assetNodeName = (assetNode.localName || assetNode.nodeName).toLowerCase();
    if (assetNodeName === 'object' || assetNodeName === 'w:object') {
      const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      const rElement = doc.createElementNS(wNamespace, "w:r");
      rElement.appendChild(assetNode);
      finalNodeToInsert = rElement;
    }

    // Replace the text run rNode with the exact imported assetNode
    rNode.parentNode?.replaceChild(finalNodeToInsert, rNode);
  }

  // Serialize and write back document.xml
  const serializer = new XMLSerializer();
  zip.file("word/document.xml", serializer.serializeToString(doc));

  // Serialize and write back document.xml.rels
  zip.file("word/_rels/document.xml.rels", serializer.serializeToString(relsDoc));

  // Ensure content types registry matches
  const contentTypesXmlText = await zip.file("[Content_Types].xml")?.async("string");
  if (contentTypesXmlText) {
    const ctDoc = parser.parseFromString(contentTypesXmlText, "application/xml");
    const typesTag = ctDoc.documentElement;
    if (typesTag) {
      // Ensure default content types are there (Namespace-agnostic / lowercase-robust matching)
      const defaults: Element[] = [];
      const allCtElements = ctDoc.getElementsByTagName("*");
      for (let i = 0; i < allCtElements.length; i++) {
        const tagLocalName = (allCtElements[i].localName || allCtElements[i].nodeName).toLowerCase();
        if (tagLocalName === 'default') {
          defaults.push(allCtElements[i]);
        }
      }
      const extensions = new Set<string>();
      for (let i = 0; i < defaults.length; i++) {
        extensions.add(defaults[i].getAttribute("Extension")?.toLowerCase() || "");
      }

      const reqExtensions = [
        { ext: "bin", mime: "application/vnd.openxmlformats-officedocument.oleObject" },
        { ext: "wmf", mime: "image/x-wmf" },
        { ext: "emf", mime: "image/x-emf" }
      ];

      const ctNs = typesTag.namespaceURI || "http://schemas.openxmlformats.org/package/2006/content-types";
      reqExtensions.forEach(({ ext, mime }) => {
        if (!extensions.has(ext)) {
          const d = ctDoc.createElementNS(ctNs, "Default");
          d.setAttribute("Extension", ext);
          d.setAttribute("ContentType", mime);
          typesTag.appendChild(d);
        }
      });

      zip.file("[Content_Types].xml", serializer.serializeToString(ctDoc));
    }
  }

  return await zip.generateAsync({ type: "blob" });
};

// --- LaTeX to docx math structure compiler helpers for export ---

const tokenize = (latex: string): Token[] => {
  const tokens: Token[] = [];
  let i = 0;
  while (i < latex.length) {
    const char = latex[i];
    if (char === '{') {
      tokens.push({ type: 'LBRACE', value: '{' });
      i++;
    } else if (char === '}') {
      tokens.push({ type: 'RBRACE', value: '}' });
      i++;
    } else if (char === '[') {
      tokens.push({ type: 'LBRACKET', value: '[' });
      i++;
    } else if (char === ']') {
      tokens.push({ type: 'RBRACKET', value: ']' });
      i++;
    } else if (char === '(') {
      tokens.push({ type: 'LPAREN', value: '(' });
      i++;
    } else if (char === ')') {
      tokens.push({ type: 'RPAREN', value: ')' });
      i++;
    } else if (char === '_') {
      tokens.push({ type: 'SUB', value: '_' });
      i++;
    } else if (char === '^') {
      tokens.push({ type: 'SUP', value: '^' });
      i++;
    } else if (char === '\\') {
      let val = '\\';
      i++;
      if (i < latex.length) {
        if (/[a-zA-Z]/.test(latex[i])) {
          while (i < latex.length && /[a-zA-Z]/.test(latex[i])) {
            val += latex[i];
            i++;
          }
        } else {
          val += latex[i];
          i++;
        }
      }
      tokens.push({ type: 'CMD', value: val });
    } else if (/\s/.test(char)) {
      i++;
    } else {
      let textVal = '';
      while (i < latex.length && !(/[{}[\]()_^\\\s]/.test(latex[i]))) {
        textVal += latex[i];
        i++;
      }
      tokens.push({ type: 'TEXT', value: textVal });
    }
  }
  return tokens;
};

const compileTokens = (tokens: Token[]): MathComponent[] => {
  let index = 0;

  const parseExpression = (): MathComponent[] => {
    const components: MathComponent[] = [];
    while (index < tokens.length) {
      const token = tokens[index];
      if (token.type === 'RBRACE' || token.type === 'RBRACKET' || token.type === 'RPAREN') {
        break;
      }

      const nextComponent = parseNode();
      if (nextComponent) {
        components.push(nextComponent);
      }
    }
    return components;
  };

  const parseNode = (): MathComponent | null => {
    if (index >= tokens.length) return null;
    const token = tokens[index];

    if (token.type === 'TEXT') {
      index++;
      return applyScripts(new MathRun(token.value));
    }

    if (token.type === 'CMD') {
      const cmd = token.value;
      index++;

      if (cmd === '\\frac') {
        const num = parseBracedGroup();
        const den = parseBracedGroup();
        return applyScripts(new MathFraction({ numerator: num, denominator: den }));
      }

      if (cmd === '\\sqrt') {
        let deg: MathComponent[] | undefined = undefined;
        if (index < tokens.length && tokens[index].type === 'LBRACKET') {
          deg = parseBracketedGroup();
        }
        const base = parseBracedGroup();
        return applyScripts(new MathRadical({ children: base, degree: deg }));
      }

      if (cmd === '\\left') {
        if (index < tokens.length) {
          const next = tokens[index];
          index++;
          const inner = parseExpression();
          if (index < tokens.length && tokens[index].value === '\\right') {
            index++;
            if (index < tokens.length) index++;
          }
          if (next.value === '(') return applyScripts(new MathRoundBrackets({ children: inner }));
          if (next.value === '[') return applyScripts(new MathSquareBrackets({ children: inner }));
          if (next.value === '{') return applyScripts(new MathCurlyBrackets({ children: inner }));
          return applyScripts(new MathRoundBrackets({ children: inner }));
        }
      }

      let symbol = cmd;
      const GreekMap: { [key: string]: string } = {
        '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ', '\\epsilon': 'ε',
        '\\zeta': 'ζ', '\\eta': 'η', '\\theta': 'θ', '\\iota': 'ι', '\\kappa': 'κ',
        '\\lambda': 'λ', '\\mu': 'μ', '\\nu': 'ν', '\\xi': 'ξ', '\\pi': 'π',
        '\\rho': 'ρ', '\\sigma': 'σ', '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ',
        '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
        '\\Delta': 'Δ', '\\Omega': 'Ω', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Theta': 'Θ', '\\Phi': 'Φ',
        '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷', '\\infty': '∞',
        '\\le': '≤', '\\ge': '≥', '\\neq': '≠', '\\approx': '≈', '\\sim': '∼',
        '\\in': 'in', '\\subset': '⊂', '\\supset': '⊃', '\\cap': '∩', '\\cup': '∪'
      };
      if (GreekMap[cmd]) {
        symbol = GreekMap[cmd];
      } else {
        symbol = cmd.replace('\\', '');
      }

      return applyScripts(new MathRun(symbol));
    }

    if (token.type === 'LPAREN') {
      index++;
      const inner = parseExpression();
      if (index < tokens.length && tokens[index].type === 'RPAREN') index++;
      return applyScripts(new MathRoundBrackets({ children: inner }));
    }

    if (token.type === 'LBRACKET') {
      index++;
      const inner = parseExpression();
      if (index < tokens.length && tokens[index].type === 'RBRACKET') index++;
      return applyScripts(new MathSquareBrackets({ children: inner }));
    }

    if (token.type === 'LBRACE') {
      index++;
      const inner = parseExpression();
      if (index < tokens.length && tokens[index].type === 'RBRACE') index++;
      return applyScripts(new MathCurlyBrackets({ children: inner }));
    }

    index++;
    return null;
  };

  const parseBracedGroup = (): MathComponent[] => {
    if (index < tokens.length && tokens[index].type === 'LBRACE') {
      index++;
      const res = parseExpression();
      if (index < tokens.length && tokens[index].type === 'RBRACE') {
        index++;
      }
      return res;
    }
    const single = parseNode();
    return single ? [single] : [];
  };

  const parseBracketedGroup = (): MathComponent[] => {
    if (index < tokens.length && tokens[index].type === 'LBRACKET') {
      index++;
      const res = parseExpression();
      if (index < tokens.length && tokens[index].type === 'RBRACKET') {
        index++;
      }
      return res;
    }
    return [];
  };

  const applyScripts = (base: MathComponent): MathComponent => {
    if (index < tokens.length) {
      if (tokens[index].type === 'SUB') {
        index++;
        const sub = parseBracedGroup();
        if (index < tokens.length && tokens[index].type === 'SUP') {
          index++;
          const sup = parseBracedGroup();
          return new MathSubSuperScript({ children: [base], subScript: sub, superScript: sup });
        }
        return new MathSubScript({ children: [base], subScript: sub });
      } else if (tokens[index].type === 'SUP') {
        index++;
        const sup = parseBracedGroup();
        if (index < tokens.length && tokens[index].type === 'SUB') {
          index++;
          const sub = parseBracedGroup();
          return new MathSubSuperScript({ children: [base], subScript: sub, superScript: sup });
        }
        return new MathSuperScript({ children: [base], superScript: sup });
      }
    }
    return base;
  };

  return parseExpression();
};

const parseLatexToDocxMath = (latex: string): MathComponent[] => {
  try {
    const tokens = tokenize(latex);
    return compileTokens(tokens);
  } catch (err) {
    console.error("Failed to parse LaTeX to docx math structure, falling back to MathRun", err, latex);
    return [new MathRun(latex)];
  }
};

// --- Components ---

const SidebarItem = ({ icon: Icon, label, active, onClick }: any) => (
  <div 
    onClick={onClick}
    className={`mcmix-sidebar-item ${active ? 'active' : ''}`}
  >
    <Icon size={20} />
    <span className="font-medium">{label}</span>
  </div>
);

const StatCard = ({ label, value, icon: Icon, color }: any) => (
  <div className="card flex items-center gap-4">
    <div className={`p-3 rounded-lg ${color}`}>
      <Icon size={24} className="text-white" />
    </div>
    <div>
      <p className="text-sm text-slate-500 font-medium">{label}</p>
      <p className="text-2xl font-bold">{value}</p>
    </div>
  </div>
);

export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [questions, setQuestions] = useState<Question[]>([]);
  const [stats, setStats] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterTopic, setFilterTopic] = useState('Tất cả');
  const [isMixing, setIsMixing] = useState(false);
  const [mixProgress, setMixProgress] = useState(0);
  const [generatedExams, setGeneratedExams] = useState<GeneratedExam[]>([]);
  const [selectedIds, setSelectedIds] = useState<number[]>([]);

  // Modal State
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalType, setModalType] = useState<'edit' | 'delete' | 'add' | 'bulk-delete'>('edit');
  const [currentQuestion, setCurrentQuestion] = useState<Partial<Question>>({});

  // AI State
  const [aiPrompt, setAiPrompt] = useState('');
  const [isAiGenerating, setIsAiGenerating] = useState(false);
  const [aiResult, setAiResult] = useState<Question[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const quillRef = useRef<any>(null);

  // Structure State
  const [structure, setStructure] = useState<ExamStructure>({
    total_questions: 28,
    part1_count: 18,
    part2_count: 4,
    part3_count: 6,
    difficulty_ratios: { 'Dễ': 40, 'Trung bình': 40, 'Khó': 20 },
    topic_distribution: {}
  });
  const [examCodes, setExamCodes] = useState<string[]>(['101', '102', '103', '104']);
  const [headerConfig, setHeaderConfig] = useState({
    schoolName: "BỘ GIÁO DỤC VÀ ĐÀO TẠO",
    examName: "KỲ THI TỐT NGHIỆP THPT TỪ NĂM 2025",
    examType: "ĐỀ THI THAM KHẢO",
    subject: "MÔN: ĐỊA LÍ",
    time: "50 phút",
    numPages: "04 trang"
  });

  const [partLabels, setPartLabels] = useState({
    part1: "PHẦN I. Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu chỉ chọn một phương án.",
    part2: "PHẦN II. Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu có các ý a), b), c), d; thí sinh đánh dấu Đúng hoặc Sai cho từng ý.",
    part3: "PHẦN III. Thí sinh trả lời từ câu 1 đến câu 6. Mỗi câu trả lời đúng được 0,25 điểm."
  });

  useEffect(() => {
    setPartLabels({
      part1: `PHẦN I. Thí sinh trả lời từ câu 1 đến câu ${structure.part1_count}. Mỗi câu chỉ chọn một phương án.`,
      part2: `PHẦN II. Thí sinh trả lời từ câu 1 đến câu ${structure.part2_count}. Mỗi câu có các ý a), b), c), d; thí sinh đánh dấu Đúng hoặc Sai cho từng ý.`,
      part3: `PHẦN III. Thí sinh trả lời từ câu 1 đến câu ${structure.part3_count}. Mỗi câu trả lời đúng được 0,25 điểm.`
    });
  }, [structure.part1_count, structure.part2_count, structure.part3_count]);

  useEffect(() => {
    if (typeof (window as any).MathJax !== 'undefined' && (window as any).MathJax.typesetPromise) {
      const timer = setTimeout(() => {
        (window as any).MathJax.typesetPromise()
          .catch((err: any) => console.error("MathJax Render Error:", err));
      }, 350);
      return () => clearTimeout(timer);
    }
  }, [activeTab, questions, generatedExams, isModalOpen]);

  useEffect(() => {
    fetchQuestions();
    fetchStats();
  }, []);

  const fetchQuestions = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/questions');
      const data = await res.json();
      setQuestions(data);
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleBulkDelete = async () => {
    if (selectedIds.length === 0) return;
    setModalType('bulk-delete');
    setIsModalOpen(true);
  };

  const confirmBulkDelete = async () => {
    try {
      const res = await fetch('/api/questions/delete-bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids: selectedIds })
      });
      if (res.ok) {
        setSelectedIds([]);
        setIsModalOpen(false);
        fetchQuestions();
        fetchStats();
      }
    } catch (err) {
      console.error(err);
    }
  };

  const toggleSelectAll = () => {
    if (selectedIds.length === filteredQuestions.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredQuestions.map(q => q.id!).filter(id => id !== undefined));
    }
  };

  const toggleSelect = (id: number) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const fetchStats = async () => {
    try {
      const res = await fetch('/api/stats');
      const data = await res.json();
      setStats(data);
      
      // Update default topic distribution if empty
      if (Object.keys(structure.topic_distribution).length === 0 && data.topicStats.length > 0) {
        const dist: Record<string, number> = {};
        const countPerTopic = Math.floor(structure.total_questions / data.topicStats.length);
        data.topicStats.forEach((s: any) => {
          dist[s.subject] = countPerTopic;
        });
        setStructure(prev => ({ ...prev, topic_distribution: dist }));
      }
    } catch (err) {
      console.error(err);
    }
  };

  const downloadExcelSample = () => {
    const data = [
      {
        Loai: "SINGLE_CHOICE",
        NoiDung: "Thủ đô của Việt Nam là gì?",
        A: "Hồ Chí Minh",
        B: "Đà Nẵng",
        C: "Hà Nội",
        D: "Cần Thơ",
        DapAn: "C",
        MonHoc: "Địa lý",
        DoKho: "Dễ",
        HinhAnh: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Flag_of_Vietnam.svg/2000px-Flag_of_Vietnam.svg.png",
        LoiGiai: "Hà Nội là thủ đô của Việt Nam."
      },
      {
        Loai: "TRUE_FALSE",
        NoiDung: "Về vị trí địa lý của Việt Nam:",
        A: "Nằm ở rìa phía đông của bán đảo Đông Dương.",
        B: "Tiếp giáp với Biển Đông ở phía Tây.",
        C: "Có đường biên giới trên đất liền dài nhất với Lào.",
        D: "Nằm hoàn toàn trong vùng nội chí tuyến Bắc bán cầu.",
        DapAn: "T,F,F,T",
        MonHoc: "Địa lý",
        DoKho: "Trung bình",
        LoiGiai: "Việt Nam giáp Biển Đông ở phía Đông."
      },
      {
        Loai: "SHORT_ANSWER",
        NoiDung: "Việt Nam có bao nhiêu tỉnh thành trực thuộc trung ương?",
        DapAn: "63",
        MonHoc: "Địa lý",
        DoKho: "Dễ",
        LoiGiai: "Gồm 58 tỉnh và 5 thành phố trực thuộc trung ương."
      }
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MauCauHoi");
    XLSX.writeFile(wb, "Mau_Nhap_Cau_Hoi_3_Phan.xlsx");
  };

  const downloadWordSample = async () => {
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            children: [new TextRun({ text: "MẪU NHẬP CÂU HỎI 3 PHẦN (WORD)", bold: true, size: 28 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
          }),
          
          // Part 1
          new Paragraph({ children: [new TextRun({ text: "PHẦN 1: TRẮC NGHIỆM 1 LỰA CHỌN", bold: true })] }),
          new Paragraph({
            children: [new TextRun({ text: "Câu 1: Thủ đô của Việt Nam là gì?", bold: true })],
            spacing: { before: 200 }
          }),
          new Paragraph({ children: [new TextRun({ text: "A. Hồ Chí Minh" })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "B. Đà Nẵng" })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "C. Hà Nội" })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "D. Cần Thơ" })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "Đáp án: C", bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: "Môn học: Địa lý", italics: true })] }),
          
          // Part 2
          new Paragraph({ children: [new TextRun({ text: "PHẦN 2: TRẮC NGHIỆM ĐÚNG/SAI", bold: true })], spacing: { before: 400 } }),
          new Paragraph({
            children: [new TextRun({ text: "Câu 2: Về vị trí địa lý của Việt Nam:", bold: true })],
            spacing: { before: 200 }
          }),
          new Paragraph({ children: [new TextRun({ text: "a) Nằm ở rìa phía đông của bán đảo Đông Dương." })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "b) Tiếp giáp với Biển Đông ở phía Tây." })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "c) Có đường biên giới trên đất liền dài nhất với Lào." })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "d) Nằm hoàn toàn trong vùng nội chí tuyến Bắc bán cầu." })], indent: { left: 720 } }),
          new Paragraph({ children: [new TextRun({ text: "Đáp án: T,F,F,T", bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: "Môn học: Địa lý", italics: true })] }),

          // Part 3
          new Paragraph({ children: [new TextRun({ text: "PHẦN 3: TRẢ LỜI NGẮN", bold: true })], spacing: { before: 400 } }),
          new Paragraph({
            children: [new TextRun({ text: "Câu 3: Việt Nam có bao nhiêu tỉnh thành trực thuộc trung ương?", bold: true })],
            spacing: { before: 200 }
          }),
          new Paragraph({ children: [new TextRun({ text: "Đáp án: 63", bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: "Môn học: Địa lý", italics: true })] }),
        ]
      }]
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "Mau_Nhap_Cau_Hoi_3_Phan.docx");
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      const formattedQuestions = data.map((row: any) => ({
        type: row.Loai || 'SINGLE_CHOICE',
        content: convertRawTextToHtmlWithFormulas(row.NoiDung || ''),
        option_a: convertRawTextToHtmlWithFormulas(row.A || ''),
        option_b: convertRawTextToHtmlWithFormulas(row.B || ''),
        option_c: convertRawTextToHtmlWithFormulas(row.C || ''),
        option_d: convertRawTextToHtmlWithFormulas(row.D || ''),
        correct_answer: row.DapAn || 'A',
        topic: row.MonHoc || row.ChuyenDe || 'Chưa phân loại',
        difficulty: row.DoKho || 'Trung bình',
        image_url: row.HinhAnh || '',
        explanation: convertRawTextToHtmlWithFormulas(row.LoiGiai || '')
      }));

      setLoading(true);
      try {
        await fetch('/api/questions/bulk', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(formattedQuestions)
        });
        fetchQuestions();
        fetchStats();
      } catch (err) {
        console.error(err);
      } finally {
        setLoading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleWordImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      let arrayBuffer = evt.target?.result as ArrayBuffer;
      try {
        // Run preprocessing on Word document to parse internal Office Math / MathType XML XML structures to LaTeX $ ... $ delimiters
        try {
          arrayBuffer = await preprocessDocx(arrayBuffer);
        } catch (mathErr) {
          console.error("Math preprocessing failed, using original document:", mathErr);
        }
        const result = await mammoth.extractRawText({ arrayBuffer });
        const text = result.value;
        
        // Simple parsing logic for Word format
        // Expected format:
        // Câu 1: Content...
        // A. Option A
        // B. Option B
        // C. Option C
        // D. Option D
        // Đáp án: A
        
        const questionBlocks = text.split(/Câu \d+[:.]/i).filter(block => block.trim() !== '');
        const parsedQuestions = questionBlocks.map(block => {
          const lines = block.split('\n').map(l => l.trim()).filter(l => l !== '');
          
          let content = '';
          let a = '', b = '', c = '', d = '', ans = '';
          let topic = 'Nhập từ Word', diff = 'Trung bình', expl = '', imgUrl = '';
          let type: QuestionType = 'SINGLE_CHOICE';
          
          let currentPart = 'content';
          
          lines.forEach(line => {
            const lowerLine = line.toLowerCase();
            if (line.startsWith('A.')) { a = line.substring(2).trim(); currentPart = 'a'; }
            else if (line.startsWith('B.')) { b = line.substring(2).trim(); currentPart = 'b'; }
            else if (line.startsWith('C.')) { c = line.substring(2).trim(); currentPart = 'c'; }
            else if (line.startsWith('D.')) { d = line.substring(2).trim(); currentPart = 'd'; }
            else if (line.startsWith('a)')) { a = line.substring(2).trim(); currentPart = 'a'; type = 'TRUE_FALSE'; }
            else if (line.startsWith('b)')) { b = line.substring(2).trim(); currentPart = 'b'; type = 'TRUE_FALSE'; }
            else if (line.startsWith('c)')) { c = line.substring(2).trim(); currentPart = 'c'; type = 'TRUE_FALSE'; }
            else if (line.startsWith('d)')) { d = line.substring(2).trim(); currentPart = 'd'; type = 'TRUE_FALSE'; }
            else if (lowerLine.startsWith('đáp án:') || lowerLine.startsWith('dap an:')) {
              ans = line.split(':')[1]?.trim() || '';
              currentPart = 'ans';
            }
            else if (lowerLine.startsWith('môn học:') || lowerLine.startsWith('mon hoc:') || lowerLine.startsWith('chuyên đề:') || lowerLine.startsWith('chuyen de:')) {
              topic = line.split(':')[1]?.trim() || 'Chưa phân loại';
              currentPart = 'topic';
            }
            else if (lowerLine.startsWith('độ khó:') || lowerLine.startsWith('do kho:')) {
              diff = line.split(':')[1]?.trim() || 'Trung bình';
              currentPart = 'diff';
            }
            else if (lowerLine.startsWith('hình ảnh:') || lowerLine.startsWith('hinh anh:') || lowerLine.startsWith('image:')) {
              imgUrl = line.split(':')[1]?.trim() || '';
            }
            else if (lowerLine.startsWith('lời giải:') || lowerLine.startsWith('loi giai:')) {
              expl = line.split(':')[1]?.trim() || '';
              currentPart = 'expl';
            }
            else {
              if (currentPart === 'content') content += ' ' + line;
              else if (currentPart === 'a') a += ' ' + line;
              else if (currentPart === 'b') b += ' ' + line;
              else if (currentPart === 'c') c += ' ' + line;
              else if (currentPart === 'd') d += ' ' + line;
              else if (currentPart === 'expl') expl += ' ' + line;
            }
          });

          // Infer type if not already set
          if (type === 'SINGLE_CHOICE' && !a && !b && !c && !d) {
            type = 'SHORT_ANSWER';
          }
          
          const finalContent = convertRawTextToHtmlWithFormulas(content.trim());
          const finalA = convertRawTextToHtmlWithFormulas(a.trim());
          const finalB = convertRawTextToHtmlWithFormulas(b.trim());
          const finalC = convertRawTextToHtmlWithFormulas(c.trim());
          const finalD = convertRawTextToHtmlWithFormulas(d.trim());
          const finalExpl = convertRawTextToHtmlWithFormulas(expl.trim());

          // Populate image_url automatically with first inline image from question's content
          let finalImgUrl = imgUrl;
          if (!finalImgUrl && finalContent) {
            const parser = new DOMParser();
            const parsedDoc = parser.parseFromString(finalContent, 'text/html');
            const firstImg = parsedDoc.querySelector('img');
            if (firstImg) {
              const src = firstImg.getAttribute('src');
              if (src && !src.includes('/api/math-assets/')) {
                finalImgUrl = src;
              }
            }
          }

          return {
            type,
            content: finalContent,
            option_a: finalA,
            option_b: finalB,
            option_c: finalC,
            option_d: finalD,
            correct_answer: ans || ((type as string) === 'TRUE_FALSE' ? 'T,T,T,T' : (type === 'SHORT_ANSWER' ? '' : 'A')),
            topic: topic,
            difficulty: diff as any,
            image_url: finalImgUrl,
            explanation: finalExpl
          };
        }).filter(q => q.content !== '');

        if (parsedQuestions.length > 0) {
          setLoading(true);
          await fetch('/api/questions/bulk', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(parsedQuestions)
          });
          fetchQuestions();
          fetchStats();
        }
      } catch (err) {
        console.error('Error parsing Word file:', err);
        // alert is discouraged, but for now we'll keep it or use a toast if we had one.
        // Actually, let's just console.error for now or use a custom modal for errors later.
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const openAddModal = () => {
    setCurrentQuestion({
      type: 'SINGLE_CHOICE',
      content: '',
      option_a: '',
      option_b: '',
      option_c: '',
      option_d: '',
      correct_answer: 'A',
      topic: 'Chưa phân loại',
      difficulty: 'Trung bình',
      explanation: ''
    });
    setModalType('add');
    setIsModalOpen(true);
  };

  const openEditModal = (q: Question) => {
    setCurrentQuestion(q);
    setModalType('edit');
    setIsModalOpen(true);
  };

  const openDeleteModal = (q: Question) => {
    setCurrentQuestion(q);
    setModalType('delete');
    setIsModalOpen(true);
  };

  const handleDelete = async () => {
    if (!currentQuestion.id) return;
    try {
      await fetch(`/api/questions/${currentQuestion.id}`, { method: 'DELETE' });
      setIsModalOpen(false);
      fetchQuestions();
      fetchStats();
    } catch (err) {
      console.error(err);
    }
  };

  const handleSave = async () => {
    const method = modalType === 'add' ? 'POST' : 'PUT';
    const url = modalType === 'add' ? '/api/questions' : `/api/questions/${currentQuestion.id}`;
    
    // Extract first image from content to populate image_url for preview
    let imageUrl = currentQuestion.image_url || '';
    if (currentQuestion.content) {
      const doc = new DOMParser().parseFromString(currentQuestion.content, 'text/html');
      const firstImg = doc.querySelector('img');
      if (firstImg) {
        imageUrl = firstImg.getAttribute('src') || '';
      }
    }

    try {
      await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...currentQuestion, image_url: imageUrl })
      });
      setIsModalOpen(false);
      fetchQuestions();
      fetchStats();
    } catch (err) {
      console.error(err);
    }
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('image', file);

    setIsUploading(true);
    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (data.url) {
        setCurrentQuestion(prev => ({ ...prev, image_url: data.url }));
      }
    } catch (err) {
      console.error('Error uploading image:', err);
    } finally {
      setIsUploading(false);
    }
  };

  const handleAiGenerate = async () => {
    if (!aiPrompt.trim()) return;
    setIsAiGenerating(true);
    try {
      const questions = await generateQuestionsWithAI(aiPrompt);
      setAiResult(questions);
    } catch (err: any) {
      console.error(err);
      alert(err.message || "Có lỗi xảy ra khi gọi AI.");
    } finally {
      setIsAiGenerating(false);
    }
  };

  const importAiQuestions = async () => {
    if (aiResult.length === 0) return;
    setLoading(true);
    try {
      await fetch('/api/questions/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(aiResult)
      });
      setAiResult([]);
      setAiPrompt('');
      fetchQuestions();
      fetchStats();
      setActiveTab('bank');
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const modules = useMemo(() => ({
    toolbar: {
      container: [
        [{ 'header': [1, 2, false] }],
        ['bold', 'italic', 'underline', 'strike', 'blockquote'],
        [{ 'list': 'ordered' }, { 'list': 'bullet' }, { 'indent': '-1' }, { 'indent': '+1' }],
        ['link', 'image', 'formula'],
        ['clean']
      ],
      handlers: {
        image: () => {
          const input = document.createElement('input');
          input.setAttribute('type', 'file');
          input.setAttribute('accept', 'image/*');
          input.click();
          input.onchange = async () => {
            const file = input.files?.[0];
            if (file) {
              setIsUploading(true);
              const formData = new FormData();
              formData.append('image', file);
              try {
                const res = await fetch('/api/upload', {
                  method: 'POST',
                  body: formData,
                });
                const data = await res.json();
                if (data.url) {
                  const quill = quillRef.current.getEditor();
                  const range = quill.getSelection();
                  quill.insertEmbed(range.index, 'image', data.url);
                }
              } catch (err) {
                console.error(err);
              } finally {
                setIsUploading(false);
              }
            }
          };
        }
      }
    },
  }), []);

  const handleMix = () => {
    setIsMixing(true);
    setMixProgress(0);
    
    // Simulate progress
    const interval = setInterval(() => {
      setMixProgress(prev => {
        if (prev >= 100) {
          clearInterval(interval);
          const exams = mixExams(questions, structure, examCodes);
          setGeneratedExams(exams);
          setIsMixing(false);
          setActiveTab('export');
          return 100;
        }
        return prev + 10;
      });
    }, 100);
  };

  const fetchImageAsBuffer = async (url: string): Promise<Uint8Array | null> => {
    try {
      if (url.startsWith('data:')) {
        const parts = url.split(',');
        if (parts.length > 1) {
          return base64ToUint8Array(parts[1]);
        }
      }
      const response = await fetch(url);
      const buffer = await response.arrayBuffer();
      return new Uint8Array(buffer);
    } catch (error) {
      console.error('Error fetching image:', error);
      return null;
    }
  };

  const fetchImageAsBase64 = async (url: string): Promise<string | null> => {
    try {
      const response = await fetch(url);
      const blob = await response.blob();
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.readAsDataURL(blob);
      });
    } catch (error) {
      console.error('Error fetching image as base64:', error);
      return null;
    }
  };
  const getAllImageUrls = (exam: GeneratedExam) => {
    const urls = new Set<string>();
    const extract = (q: Question) => {
      if (q.image_url) urls.add(q.image_url);
      
      // Parse main content for MathType images
      const doc = new DOMParser().parseFromString(ensureMathAssetsRendered(q.content || ''), 'text/html');
      doc.querySelectorAll('img').forEach(img => {
        const src = img.getAttribute('src');
        if (src) urls.add(src);
      });

      // Parse options for MathType images
      const opts = [q.option_a, q.option_b, q.option_c, q.option_d];
      opts.forEach(opt => {
        if (opt) {
          const optDoc = new DOMParser().parseFromString(ensureMathAssetsRendered(opt), 'text/html');
          optDoc.querySelectorAll('img').forEach(img => {
            const src = img.getAttribute('src');
            if (src) urls.add(src);
          });
        }
      });
    };
    exam.part1.forEach(extract);
    exam.part2.forEach(extract);
    exam.part3.forEach(extract);
    return Array.from(urls);
  };

  const renderHtmlToParagraphChildren = (html: string, imageMap: Map<string, Uint8Array> | null = null): any[] => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html || '', 'text/html');
    const children: any[] = [];

    const process = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        if (node.textContent && node.textContent.trim()) {
          children.push(new TextRun({ text: convertTexDelimiters(node.textContent) }));
        }
      } else if (node instanceof Element && node.classList.contains('ql-formula')) {
        const formula = node.getAttribute('data-value');
        if (formula) {
          try {
            const mathComponents = parseLatexToDocxMath(formula);
            children.push(new DocxMath({ children: mathComponents }));
          } catch (e) {
            console.error("Error creating docx native math", e);
            children.push(new TextRun({ text: `$${formula.trim()}$` }));
          }
        }
      } else if (node.nodeName === 'IMG') {
        const element = node as HTMLImageElement;
        const isMathAsset = element.classList.contains('inline-math-asset') || (element.getAttribute('src') || '').includes('/api/math-assets/');
        if (isMathAsset) {
          const assetId = element.getAttribute('alt') || '';
          if (assetId) {
            children.push(new TextRun({ text: `[!m:${assetId}]` }));
          }
        } else if (imageMap) {
          const src = element.getAttribute('src');
          const buffer = src ? imageMap.get(src) : null;
          if (buffer) {
            children.push(new ImageRun({
              data: buffer,
              transformation: { width: 300, height: 200 },
              type: 'png'
            }));
          }
        }
      } else {
        node.childNodes.forEach(process);
      }
    };

    process(doc.body);
    return children;
  };

  const renderQuestionContentDocx = (q: Question, prefix: string, imageMap: Map<string, Uint8Array>) => {
    const html = ensureMathAssetsRendered(q.content || '');
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const paragraphs: Paragraph[] = [];
    let currentRuns: any[] = [];
    let isFirst = true;

    const processNode = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        if (node.textContent && node.textContent.trim()) {
          if (isFirst) {
            currentRuns.push(new TextRun({ text: prefix, bold: true }));
            isFirst = false;
          }
          currentRuns.push(new TextRun({ text: convertTexDelimiters(node.textContent) }));
        }
      } else if (node instanceof Element && node.classList.contains('ql-formula')) {
        const formula = node.getAttribute('data-value');
        if (formula) {
          if (isFirst) {
            currentRuns.push(new TextRun({ text: prefix, bold: true }));
            isFirst = false;
          }
          try {
            const mathComponents = parseLatexToDocxMath(formula);
            currentRuns.push(new DocxMath({ children: mathComponents }));
          } catch (e) {
            console.error("Error creating docx native math, falling back to text run", e);
            currentRuns.push(new TextRun({ text: `$${formula.trim()}$` }));
          }
        }
      } else if (node.nodeName === 'IMG') {
        const element = node as HTMLImageElement;
        const isMathAsset = element.classList.contains('inline-math-asset') || (element.getAttribute('src') || '').includes('/api/math-assets/');
        if (isMathAsset) {
          const assetId = element.getAttribute('alt') || '';
          if (assetId) {
            if (isFirst) {
              currentRuns.push(new TextRun({ text: prefix, bold: true }));
              isFirst = false;
            }
            currentRuns.push(new TextRun({ text: `[!m:${assetId}]` }));
          }
        } else {
          const src = element.getAttribute('src');
          const buffer = src ? imageMap.get(src) : null;
          if (buffer) {
            if (isFirst) {
              currentRuns.push(new TextRun({ text: prefix, bold: true }));
              isFirst = false;
            }
            currentRuns.push(new ImageRun({
              data: buffer,
              transformation: { width: 300, height: 200 },
              type: 'png'
            }));
          }
        }
      } else if (['P', 'DIV', 'LI', 'BR'].includes(node.nodeName)) {
        if (currentRuns.length > 0) {
          paragraphs.push(new Paragraph({ children: currentRuns, spacing: { before: 100 } }));
          currentRuns = [];
        }
        node.childNodes.forEach(processNode);
        if (currentRuns.length > 0) {
          paragraphs.push(new Paragraph({ children: currentRuns, spacing: { before: 100 } }));
          currentRuns = [];
        }
      } else {
        node.childNodes.forEach(processNode);
      }
    };

    processNode(doc.body);

    if (currentRuns.length > 0) {
      paragraphs.push(new Paragraph({ children: currentRuns, spacing: { before: 100 } }));
    }

    if (paragraphs.length === 0) {
      paragraphs.push(new Paragraph({ 
        children: [new TextRun({ text: prefix, bold: true }), new TextRun({ text: stripHtml(q.content) })], 
        spacing: { before: 100 } 
      }));
    }

    // Handle legacy image_url if not already in content
    if (q.image_url) {
      const buffer = imageMap.get(q.image_url);
      const contentHasThisImg = html.includes(q.image_url);
      if (buffer && !contentHasThisImg) {
        paragraphs.push(new Paragraph({
          children: [
            new ImageRun({
              data: buffer,
              transformation: { width: 300, height: 200 },
              type: 'png'
            })
          ],
          alignment: AlignmentType.CENTER,
          spacing: { before: 100, after: 100 }
        }));
      }
    }

    return paragraphs;
  };

  const exportToWord = async (exam: GeneratedExam) => {
    // Pre-fetch all images
    const allUrls = getAllImageUrls(exam);
    const imageMap = new Map<string, Uint8Array>();
    await Promise.all(allUrls.map(async url => {
      const buffer = await fetchImageAsBuffer(url);
      if (buffer) imageMap.set(url, buffer);
    }));

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 720,
              right: 720,
              bottom: 720,
              left: 720,
            },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: `Mã đề: ${exam.code}`, bold: true }),
                  new TextRun({ text: "\t\t\t\t\t\t\t\t\t\t" }), // Tab for spacing
                  new TextRun({ text: "Trang ", bold: true }),
                  new TextRun({ children: [PageNumber.CURRENT], bold: true }),
                  new TextRun({ text: " / ", bold: true }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], bold: true }),
                ],
              }),
            ],
          }),
        },
        children: [
          // Header
          new Paragraph({
            children: [
              new TextRun({ text: headerConfig.schoolName, bold: true }),
              new TextRun({ text: `\t\t\t${headerConfig.examName}`, bold: true }),
            ],
            alignment: AlignmentType.LEFT,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: headerConfig.examType, bold: true }),
              new TextRun({ text: `\t\t\t\t${headerConfig.subject}`, bold: true }),
            ],
            alignment: AlignmentType.LEFT,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `(${headerConfig.numPages})`, italics: true }),
              new TextRun({ text: `\t\t\t\tThời gian làm bài ${headerConfig.time}, không kể thời gian phát đề`, italics: true }),
            ],
            alignment: AlignmentType.LEFT,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `Mã đề: ${exam.code}`, bold: true, size: 24 }),
            ],
            alignment: AlignmentType.RIGHT,
            spacing: { before: 200, after: 200 }
          }),

          // Part I
          new Paragraph({
            children: [
              new TextRun({ text: partLabels.part1, bold: true }),
            ],
            spacing: { before: 200, after: 200 }
          }),
          ...exam.part1.flatMap((q, idx) => {
            const paragraphs = renderQuestionContentDocx(q, `Câu ${idx + 1}. `, imageMap);

            const optA = q.shuffled_options && q.shuffled_options.length > 0
              ? q.shuffled_options.find(o => o.label === 'A')?.text || ''
              : q.option_a || '';
            const optB = q.shuffled_options && q.shuffled_options.length > 0
              ? q.shuffled_options.find(o => o.label === 'B')?.text || ''
              : q.option_b || '';
            const optC = q.shuffled_options && q.shuffled_options.length > 0
              ? q.shuffled_options.find(o => o.label === 'C')?.text || ''
              : q.option_c || '';
            const optD = q.shuffled_options && q.shuffled_options.length > 0
              ? q.shuffled_options.find(o => o.label === 'D')?.text || ''
              : q.option_d || '';

            paragraphs.push(new Paragraph({
              children: [
                new TextRun({ text: "A. " }),
                ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(optA), imageMap),
                new TextRun({ text: "\t\tB. " }),
                ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(optB), imageMap),
                new TextRun({ text: "\t\tC. " }),
                ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(optC), imageMap),
                new TextRun({ text: "\t\tD. " }),
                ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(optD), imageMap),
              ],
              indent: { left: 360 },
            }));

            return paragraphs;
          }),

          // Part II
          new Paragraph({
            children: [
              new TextRun({ text: partLabels.part2, bold: true }),
            ],
            spacing: { before: 400, after: 200 }
          }),
          ...exam.part2.flatMap((q, idx) => {
            const paragraphs = renderQuestionContentDocx(q, `Câu ${idx + 1}. `, imageMap);

            paragraphs.push(new Paragraph({ children: [new TextRun({ text: "a) " }), ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(q.option_a), imageMap)], indent: { left: 360 } }));
            paragraphs.push(new Paragraph({ children: [new TextRun({ text: "b) " }), ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(q.option_b), imageMap)], indent: { left: 360 } }));
            paragraphs.push(new Paragraph({ children: [new TextRun({ text: "c) " }), ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(q.option_c), imageMap)], indent: { left: 360 } }));
            paragraphs.push(new Paragraph({ children: [new TextRun({ text: "d) " }), ...renderHtmlToParagraphChildren(ensureMathAssetsRendered(q.option_d), imageMap)], indent: { left: 360 } }));

            return paragraphs;
          }),

          // Part III
          new Paragraph({
            children: [
              new TextRun({ text: partLabels.part3, bold: true }),
            ],
            spacing: { before: 400, after: 200 }
          }),
          ...exam.part3.flatMap((q, idx) => {
            return renderQuestionContentDocx(q, `Câu ${idx + 1}. `, imageMap);
          }),
          
          new Paragraph({
            children: [new TextRun({ text: "------------------ HẾT ------------------", bold: true })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 400 }
          })
        ],
      }],
    });

    let blob = await Packer.toBlob(doc);
    try {
      blob = await postprocessMathAssetsInDocx(blob, exam);
    } catch (postErr) {
      console.error("Error post-processing math assets in docx:", postErr);
    }
    saveAs(blob, `De_Thi_Ma_${exam.code}.docx`);
  };

  const exportAnswerKey = () => {
    if (generatedExams.length === 0) return;
    
    const data: any[] = [];
    
    // Part I
    const p1Len = generatedExams[0].part1.length;
    for (let i = 0; i < p1Len; i++) {
      const row: any = { 'Phần': 'I', 'Câu': i + 1 };
      generatedExams.forEach(exam => {
        row[exam.code] = exam.part1[i].new_correct_answer;
      });
      data.push(row);
    }

    // Part II
    const p2Len = generatedExams[0].part2.length;
    for (let i = 0; i < p2Len; i++) {
      const row: any = { 'Phần': 'II', 'Câu': i + 1 };
      generatedExams.forEach(exam => {
        row[exam.code] = exam.part2[i].correct_answer;
      });
      data.push(row);
    }

    // Part III
    const p3Len = generatedExams[0].part3.length;
    for (let i = 0; i < p3Len; i++) {
      const row: any = { 'Phần': 'III', 'Câu': i + 1 };
      generatedExams.forEach(exam => {
        row[exam.code] = exam.part3[i].correct_answer;
      });
      data.push(row);
    }

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DapAn");
    XLSX.writeFile(wb, "Bang_Dap_An_Tong_Hop.xlsx");
  };

  const removeAccents = (str: string) => {
    return str.normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/đ/g, 'd').replace(/Đ/g, 'D');
  };

  const ensureMathAssetsRendered = (html: string): string => {
    if (!html) return '';
    let result = html;
    
    // Replace old style with $ (e.g. [!m:$mathtype_1_om_1$])
    result = result.replace(/\[!m:\$(mathtype_[^\$]+?)\$\]/g, (_, assetId) => {
      return `<img class="inline-math-asset" src="/api/math-assets/${assetId}/image" alt="${assetId}" style="max-height: 2.2em; vertical-align: middle; display: inline-block; margin: 0 4px;" referrerPolicy="no-referrer" />`;
    });
    
    // Replace new style without $ (e.g. [!m:mathtype_1_om_1])
    result = result.replace(/\[!m:(mathtype_[^\]]+?)\]/g, (_, assetId) => {
      return `<img class="inline-math-asset" src="/api/math-assets/${assetId}/image" alt="${assetId}" style="max-height: 2.2em; vertical-align: middle; display: inline-block; margin: 0 4px;" referrerPolicy="no-referrer" />`;
    });
    
    return result;
  };

  const convertRawTextToHtmlWithFormulas = (text: string): string => {
    if (!text) return '';
    
    // First escape standard HTML entities to preserve text formatting inside HTML
    let escaped = text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    // 1. Double dollars: $$math$$ -> <span class="ql-formula" data-value="math">\(math\)</span>
    escaped = escaped.replace(/\$\$(.*?)\$\$/gs, (_, formula) => {
      const trimmed = formula.trim();
      return `<span class="ql-formula" data-value="${trimmed}">\\(${trimmed}\\)</span>`;
    });

    // 2. Bracket display math: \[math\] -> <span class="ql-formula" data-value="math">\[math\]</span>
    escaped = escaped.replace(/\\\[(.*?)\\\]/gs, (_, formula) => {
      const trimmed = formula.trim();
      return `<span class="ql-formula" data-value="${trimmed}">\\[${trimmed}\\]</span>`;
    });

    // 3. Parenthesis inline math: \(math\) -> <span class="ql-formula" data-value="math">\(math\)</span>
    escaped = escaped.replace(/\\\((.*?)\\\)/gs, (_, formula) => {
      const trimmed = formula.trim();
      return `<span class="ql-formula" data-value="${trimmed}">\\(${trimmed}\\)</span>`;
    });

    // 4. Single dollars math: $math$ -> <span class="ql-formula" data-value="math">\(math\)</span>
    escaped = escaped.replace(/\$([^\$]+?)\$/g, (_, formula) => {
      const trimmed = formula.trim();
      return `<span class="ql-formula" data-value="${trimmed}">\\(${trimmed}\\)</span>`;
    });

    // 5. MathType tracker placeholder replacement
    escaped = ensureMathAssetsRendered(escaped);

    // 6. Inline images replacement (replacing the placeholder with real img tag AFTER escaping)
    escaped = escaped.replace(/\[!img:(data:image\/[^\]]+?)\]/g, (_, dataUrl) => {
      return `<img src="${dataUrl}" class="inline-block max-w-full my-2 align-middle" style="max-height: 12rem;" />`;
    });

    return escaped;
  };

  const convertTexDelimiters = (text: string) => {
    if (!text) return '';
    // Preserve or convert standard TeX structures to MathType-compatible standard $ ... $ and $$ ... $$
    let result = text;
    // Cover the case of \( ... \) -> $ ... $
    result = result.replace(/\\\((.*?)\\\)/g, '$$$1$');
    // Cover the case of \[ ... \] -> $$ ... $$
    result = result.replace(/\\\[(.*?)\\\]/g, '$$$$$1$$$$');
    return result;
  };

  const stripHtml = (html: string) => {
    if (!html) return '';
    const cleaned = ensureMathAssetsRendered(html);
    const doc = new DOMParser().parseFromString(cleaned, 'text/html');
    
    // Convert Quill / KaTeX formulas to text: $formula$
    doc.querySelectorAll('.ql-formula').forEach(el => {
      const formula = el.getAttribute('data-value');
      if (formula) {
        el.textContent = ` $${formula.trim()}$ `;
      }
    });

    return convertTexDelimiters(doc.body.textContent || "");
  };

  const renderQuestionContentPdf = (doc: jsPDF, q: Question, prefix: string, imageMap: Map<string, string>, startY: number, pageWidth: number) => {
    const html = ensureMathAssetsRendered(q.content || '');
    const parser = new DOMParser();
    const htmlDoc = parser.parseFromString(html, 'text/html');
    let y = startY;
    let isFirst = true;

    const processNode = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        let text = node.textContent || '';
        if (text.trim()) {
          if (isFirst) {
            doc.setFont("helvetica", "bold");
            doc.text(prefix, 20, y);
            doc.setFont("helvetica", "normal");
            isFirst = false;
          }
          const lines = doc.splitTextToSize(removeAccents(text), pageWidth - 40);
          doc.text(lines, 35, y);
          y += (lines.length * 6) + 2;
        }
      } else if (node.nodeName === 'IMG') {
        const src = (node as HTMLImageElement).getAttribute('src');
        const base64 = src ? imageMap.get(src) : null;
        if (base64) {
          if (isFirst) {
            doc.setFont("helvetica", "bold");
            doc.text(prefix, 20, y);
            doc.setFont("helvetica", "normal");
            isFirst = false;
          }
          if (y > 230) { doc.addPage(); y = 20; }
          try {
            doc.addImage(base64, 'PNG', 40, y, 60, 40);
            y += 45;
          } catch (e) {
            console.error('Error adding image to PDF:', e);
          }
        }
      } else if (['P', 'DIV', 'LI', 'BR'].includes(node.nodeName)) {
        node.childNodes.forEach(processNode);
      } else {
        node.childNodes.forEach(processNode);
      }
    };

    processNode(htmlDoc.body);
    
    // Handle legacy image_url if not already in content
    if (q.image_url) {
      const base64 = imageMap.get(q.image_url);
      const contentHasThisImg = html.includes(q.image_url);
      if (base64 && !contentHasThisImg) {
        if (y > 230) { doc.addPage(); y = 20; }
        try {
          doc.addImage(base64, 'PNG', 40, y, 60, 40);
          y += 45;
        } catch (e) {
          console.error('Error adding legacy image to PDF:', e);
        }
      }
    }

    return y;
  };

  const exportToPDF = async (exam: GeneratedExam) => {
    // Pre-fetch all images as base64
    const allUrls = getAllImageUrls(exam);
    const imageMap = new Map<string, string>();
    await Promise.all(allUrls.map(async url => {
      const base64 = await fetchImageAsBase64(url);
      if (base64) imageMap.set(url, base64);
    }));

    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    
    // Header
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text(removeAccents(headerConfig.schoolName), 20, 20);
    doc.text(removeAccents(headerConfig.examName), pageWidth - 20, 20, { align: "right" });
    doc.text(removeAccents(headerConfig.examType), 20, 28);
    doc.text(removeAccents(headerConfig.subject), pageWidth - 20, 28, { align: "right" });
    
    doc.setFont("helvetica", "italic");
    doc.setFontSize(10);
    doc.text(`(${removeAccents(headerConfig.numPages)})`, 20, 36);
    doc.text(`Thoi gian lam bai ${removeAccents(headerConfig.time)}`, pageWidth - 20, 36, { align: "right" });
    
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text(`Ma de: ${exam.code}`, pageWidth - 20, 48, { align: "right" });
    
    let y = 60;

    // Part I
    doc.setFontSize(11);
    doc.text(removeAccents(partLabels.part1), 20, y);
    y += 10;
    
    for (let i = 0; i < exam.part1.length; i++) {
      const q = exam.part1[i];
      if (y > 250) { doc.addPage(); y = 20; }
      y = renderQuestionContentPdf(doc, q, `Cau ${i + 1}: `, imageMap, y, pageWidth);
      
      const optA = q.shuffled_options && q.shuffled_options.length > 0
        ? q.shuffled_options.find(o => o.label === 'A')?.text || ''
        : q.option_a || '';
      const optB = q.shuffled_options && q.shuffled_options.length > 0
        ? q.shuffled_options.find(o => o.label === 'B')?.text || ''
        : q.option_b || '';
      const optC = q.shuffled_options && q.shuffled_options.length > 0
        ? q.shuffled_options.find(o => o.label === 'C')?.text || ''
        : q.option_c || '';
      const optD = q.shuffled_options && q.shuffled_options.length > 0
        ? q.shuffled_options.find(o => o.label === 'D')?.text || ''
        : q.option_d || '';

      const options = `A. ${removeAccents(stripHtml(optA))}   B. ${removeAccents(stripHtml(optB))}   C. ${removeAccents(stripHtml(optC))}   D. ${removeAccents(stripHtml(optD))}`;
      doc.text(options, 35, y);
      y += 10;
    }

    // Part II
    if (y > 250) { doc.addPage(); y = 20; }
    doc.setFont("helvetica", "bold");
    doc.text(removeAccents(partLabels.part2), 20, y);
    y += 10;
    
    for (let i = 0; i < exam.part2.length; i++) {
      const q = exam.part2[i];
      if (y > 230) { doc.addPage(); y = 20; }
      y = renderQuestionContentPdf(doc, q, `Cau ${i + 1}: `, imageMap, y, pageWidth);
      
      doc.text(`a) ${removeAccents(stripHtml(q.option_a))}`, 40, y); y += 6;
      doc.text(`b) ${removeAccents(stripHtml(q.option_b))}`, 40, y); y += 6;
      doc.text(`c) ${removeAccents(stripHtml(q.option_c))}`, 40, y); y += 6;
      doc.text(`d) ${removeAccents(stripHtml(q.option_d))}`, 40, y); y += 8;
    }

    // Part III
    if (y > 250) { doc.addPage(); y = 20; }
    doc.setFont("helvetica", "bold");
    doc.text(removeAccents(partLabels.part3), 20, y);
    y += 10;
    
    for (let i = 0; i < exam.part3.length; i++) {
      const q = exam.part3[i];
      if (y > 250) { doc.addPage(); y = 20; }
      y = renderQuestionContentPdf(doc, q, `Cau ${i + 1}: `, imageMap, y, pageWidth);
      y += 6;
    }

    // Footer
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.text(`Ma de: ${exam.code}`, 20, doc.internal.pageSize.getHeight() - 10);
      doc.text(`Trang ${i} / ${totalPages}`, pageWidth - 20, doc.internal.pageSize.getHeight() - 10, { align: "right" });
    }

    doc.save(`De_Thi_Ma_${exam.code}.pdf`);
  };


  const filteredQuestions = questions.filter(q => 
    (q.content.toLowerCase().includes(searchTerm.toLowerCase()) || 
     q.topic.toLowerCase().includes(searchTerm.toLowerCase())) &&
    (filterTopic === 'Tất cả' || q.topic === filterTopic)
  );

  const topics = stats ? ['Tất cả', ...stats.topicStats.map((s: any) => s.subject)] : ['Tất cả'];

  return (
    <div className="flex h-screen bg-slate-50 overflow-hidden">
      {/* Sidebar */}
      <div className="w-64 bg-white border-r border-slate-200 flex flex-col p-4">
        <div className="flex items-center gap-2 px-2 mb-8">
          <div className="w-10 h-10 bg-primary rounded-xl flex items-center justify-center text-white font-bold text-xl">M</div>
          <div>
            <h1 className="font-bold text-slate-800 leading-tight">McMix Pro</h1>
            <p className="text-[10px] text-slate-500 uppercase tracking-wider font-bold">TN THPT 2025</p>
          </div>
        </div>

        <nav className="flex-1 space-y-1">
          <SidebarItem icon={LayoutDashboard} label="Trang chủ" active={activeTab === 'dashboard'} onClick={() => setActiveTab('dashboard')} />
          <SidebarItem icon={Database} label="Ngân hàng câu hỏi" active={activeTab === 'bank'} onClick={() => setActiveTab('bank')} />
          <SidebarItem icon={Settings} label="Cấu trúc đề thi" active={activeTab === 'structure'} onClick={() => setActiveTab('structure')} />
          <SidebarItem icon={Shuffle} label="Trộn đề" active={activeTab === 'mix'} onClick={() => setActiveTab('mix')} />
          <SidebarItem icon={Sparkles} label="Trợ lý AI" active={activeTab === 'ai'} onClick={() => setActiveTab('ai')} />
          <SidebarItem icon={Download} label="Xuất đề" active={activeTab === 'export'} onClick={() => setActiveTab('export')} />
        </nav>

        <div className="mt-auto p-4 bg-slate-50 rounded-xl border border-slate-200">
          <div className="flex items-center gap-3 mb-2">
            <div className="w-2 h-2 rounded-full bg-accent animate-pulse"></div>
            <span className="text-xs font-bold text-slate-600">Hệ thống sẵn sàng</span>
          </div>
          <p className="text-[10px] text-slate-400">Phiên bản 2.5.0 (2025 Edition)</p>
        </div>
      </div>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto p-8">
        <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8"
            >
              <header className="flex justify-between items-end">
                <div>
                  <h2 className="text-3xl font-bold text-slate-800">Tổng quan hệ thống</h2>
                  <p className="text-slate-500">Chào mừng bạn trở lại, giáo viên.</p>
                </div>
                <button onClick={() => setActiveTab('bank')} className="btn-primary flex items-center gap-2">
                  <Plus size={18} /> Thêm câu hỏi mới
                </button>
              </header>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                <StatCard label="Tổng số câu hỏi" value={stats?.total || 0} icon={Database} color="bg-blue-600" />
                <StatCard label="Độ khó: Dễ" value={stats?.difficultyStats.find((s: any) => s.difficulty === 'Dễ')?.count || 0} icon={CheckCircle2} color="bg-emerald-500" />
                <StatCard label="Độ khó: Trung bình" value={stats?.difficultyStats.find((s: any) => s.difficulty === 'Trung bình')?.count || 0} icon={AlertCircle} color="bg-amber-500" />
                <StatCard label="Độ khó: Khó" value={stats?.difficultyStats.find((s: any) => s.difficulty === 'Khó')?.count || 0} icon={AlertCircle} color="bg-rose-500" />
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                <div className="card">
                  <h3 className="font-bold text-lg mb-4 flex items-center gap-2">
                    <FileText className="text-primary" size={20} /> Môn học phổ biến
                  </h3>
                  <div className="space-y-4">
                    {stats?.topicStats.slice(0, 5).map((topic: any, idx: number) => (
                      <div key={idx} className="flex items-center justify-between">
                        <span className="text-slate-600 font-medium">{topic.subject}</span>
                        <div className="flex items-center gap-3 flex-1 mx-4">
                          <div className="h-2 bg-slate-100 rounded-full flex-1 overflow-hidden">
                            <div 
                              className="h-full bg-primary" 
                              style={{ width: `${(topic.count / stats.total) * 100}%` }}
                            ></div>
                          </div>
                          <span className="text-xs font-bold text-slate-400 w-8">{topic.count}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="card bg-primary text-white border-none relative overflow-hidden">
                  <div className="relative z-10">
                    <h3 className="font-bold text-xl mb-2">Bắt đầu trộn đề ngay</h3>
                    <p className="text-blue-100 mb-6 text-sm">Thiết lập cấu trúc và tạo hàng trăm mã đề chỉ trong vài giây.</p>
                    <button onClick={() => setActiveTab('mix')} className="bg-white text-primary px-6 py-3 rounded-xl font-bold hover:bg-blue-50 transition-colors">
                      Trộn đề ngay
                    </button>
                  </div>
                  <Shuffle className="absolute -right-8 -bottom-8 text-white/10" size={200} />
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'bank' && (
            <motion.div 
              key="bank"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-6"
            >
              <header className="flex justify-between items-center">
                <div>
                  <h2 className="text-2xl font-bold text-slate-800">Ngân hàng câu hỏi</h2>
                  <div className="flex gap-2 mt-1">
                    <button onClick={downloadExcelSample} className="text-[10px] font-bold text-emerald-600 hover:underline">Tải file mẫu Excel</button>
                    <span className="text-slate-300">|</span>
                    <button onClick={downloadWordSample} className="text-[10px] font-bold text-blue-600 hover:underline">Tải file mẫu Word</button>
                  </div>
                </div>
                <div className="flex gap-3">
                  {selectedIds.length > 0 && (
                    <button 
                      onClick={handleBulkDelete}
                      className="btn-secondary text-rose-600 border-rose-200 hover:bg-rose-50 flex items-center gap-2"
                    >
                      <Trash2 size={18} /> Xóa đã chọn ({selectedIds.length})
                    </button>
                  )}
                  <label className="btn-secondary flex items-center gap-2 cursor-pointer">
                    <Upload size={18} /> Nhập từ Excel
                    <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleExcelImport} />
                  </label>
                  <label className="btn-secondary flex items-center gap-2 cursor-pointer">
                    <FileText size={18} /> Nhập từ Word
                    <input type="file" accept=".docx" className="hidden" onChange={handleWordImport} />
                  </label>
                  <button onClick={openAddModal} className="btn-primary flex items-center gap-2">
                    <Plus size={18} /> Thêm thủ công
                  </button>
                </div>
              </header>

              <div className="bg-blue-50/60 border border-blue-100 rounded-2xl p-4 flex gap-4 text-sm text-slate-700">
                <div className="w-10 h-10 rounded-xl bg-blue-100 flex items-center justify-center flex-shrink-0 text-blue-600">
                  <Sparkles size={20} />
                </div>
                <div>
                  <h4 className="font-bold text-slate-800 mb-1 text-xs sm:text-sm">Mẹo hiển thị công thức môn Toán (MathType, Equation, LaTeX)</h4>
                  <p className="leading-relaxed text-slate-600 text-xs text-[11px] sm:text-xs">
                    Để các công thức toán học hiển thị trực quan dạng Vector sắc nét, hệ thống tự động quét và vẽ các định dạng <strong>MathML</strong> và <strong>LaTeX/TeX</strong>.
                  </p>
                  <p className="leading-relaxed text-slate-600 text-xs text-[11px] sm:text-xs mt-1">
                    • <strong>Đề thi từ Word:</strong> Thầy cô nên sử dụng tính năng <strong>Convert Equations</strong> hoặc <strong>Toggle TeX</strong> trong công cụ MathType (Word) để đổi toàn bộ công thức sang dạng ký hiệu LaTeX (như <code className="bg-blue-100 px-1 py-0.5 rounded border border-blue-200 font-mono text-[10px] text-blue-700">$x^2 - 2x + 1 = 0$</code>) trước khi lưu và upload.
                  </p>
                  <p className="leading-relaxed text-slate-600 text-xs text-[11px] sm:text-xs mt-0.5">
                    • <strong>Nhập trực tiếp:</strong> Sử dụng phím công thức <code className="bg-blue-100 px-1 py-0.5 rounded border border-blue-200 font-bold text-[10px] text-blue-700">Formula</code> trong khung soạn thảo để dễ dàng nhập phương trình toán học chuẩn xác nhất.
                  </p>
                </div>
              </div>

              <div className="card p-4 flex flex-wrap gap-4 items-center">
                <div className="relative flex-1 min-w-[300px]">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                  <input 
                    type="text" 
                    placeholder="Tìm kiếm nội dung hoặc môn học..." 
                    className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary/20"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                </div>
                <div className="flex items-center gap-2">
                  <Filter size={18} className="text-slate-400" />
                  <select 
                    className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none"
                    value={filterTopic}
                    onChange={(e) => setFilterTopic(e.target.value)}
                  >
                    {topics.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>

              <div className="card p-0 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[1000px]">
                  <thead>
                    <tr className="bg-slate-50 border-bottom border-slate-200">
                      <th className="px-6 py-4 w-10">
                        <input 
                          type="checkbox" 
                          className="rounded border-slate-300 text-primary focus:ring-primary"
                          checked={filteredQuestions.length > 0 && selectedIds.length === filteredQuestions.length}
                          onChange={toggleSelectAll}
                        />
                      </th>
                      <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">ID</th>
                      <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Loại</th>
                      <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Nội dung</th>
                      <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Môn học</th>
                      <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Độ khó</th>
                      <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider sticky right-0 bg-slate-50 shadow-[-10px_0_15px_-3px_rgba(0,0,0,0.05)]">Thao tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {loading ? (
                      <tr>
                        <td colSpan={5} className="px-6 py-12 text-center">
                          <Loader2 className="animate-spin mx-auto text-primary mb-2" />
                          <p className="text-slate-400">Đang tải dữ liệu...</p>
                        </td>
                      </tr>
                    ) : filteredQuestions.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-6 py-12 text-center text-slate-400">
                          Không tìm thấy câu hỏi nào.
                        </td>
                      </tr>
                    ) : filteredQuestions.map((q) => (
                      <tr key={q.id} className={`hover:bg-slate-50 transition-colors group ${selectedIds.includes(q.id!) ? 'bg-blue-50/50' : ''}`}>
                        <td className="px-6 py-4">
                          <input 
                            type="checkbox" 
                            className="rounded border-slate-300 text-primary focus:ring-primary"
                            checked={selectedIds.includes(q.id!)}
                            onChange={() => toggleSelect(q.id!)}
                          />
                        </td>
                        <td className="px-6 py-4 text-sm font-mono text-slate-400">#{q.id}</td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-bold px-2 py-1 rounded-md uppercase ${
                            q.type === 'SINGLE_CHOICE' ? 'bg-blue-100 text-blue-700' :
                            q.type === 'TRUE_FALSE' ? 'bg-purple-100 text-purple-700' :
                            'bg-orange-100 text-orange-700'
                          }`}>
                            {q.type === 'SINGLE_CHOICE' ? '1 Lựa chọn' :
                             q.type === 'TRUE_FALSE' ? 'Đúng/Sai' : 'Trả lời ngắn'}
                          </span>
                        </td>
                        <td className="px-6 py-4 max-w-md">
                          <div className="flex items-center gap-3">
                            {q.image_url && (
                              <div className="w-10 h-10 rounded bg-slate-100 flex-shrink-0 overflow-hidden border border-slate-200">
                                <img 
                                  src={q.image_url} 
                                  alt="" 
                                  className="w-full h-full object-cover"
                                  onError={(e) => (e.currentTarget.style.display = 'none')}
                                />
                              </div>
                            )}
                            <div 
                              className="text-sm font-medium text-slate-700 line-clamp-2 max-h-12 overflow-hidden select-text ql-editor-mini" 
                              dangerouslySetInnerHTML={{ __html: ensureMathAssetsRendered(q.content) }} 
                            />
                          </div>
                        </td>
                        <td className="px-6 py-4">
                          <span className="text-xs font-bold px-2 py-1 bg-slate-100 text-slate-600 rounded-md">{q.topic}</span>
                        </td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-bold px-2 py-1 rounded-md uppercase ${
                            q.difficulty === 'Dễ' ? 'bg-emerald-100 text-emerald-700' :
                            q.difficulty === 'Trung bình' ? 'bg-amber-100 text-amber-700' :
                            'bg-rose-100 text-rose-700'
                          }`}>
                            {q.difficulty}
                          </span>
                        </td>
                        <td className="px-6 py-4 sticky right-0 bg-white group-hover:bg-slate-50 shadow-[-10px_0_15px_-3px_rgba(0,0,0,0.05)]">
                          <div className="flex gap-2">
                            <button 
                              onClick={() => openEditModal(q)}
                              className="p-2 text-slate-400 hover:text-primary hover:bg-blue-50 rounded-lg transition-colors"
                            >
                              <Edit size={16} />
                            </button>
                            <button 
                              onClick={() => openDeleteModal(q)}
                              className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'structure' && (
            <motion.div 
              key="structure"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="max-w-4xl mx-auto space-y-8"
            >
              <header>
                <h2 className="text-2xl font-bold text-slate-800">Thiết lập cấu trúc đề thi</h2>
                <p className="text-slate-500">Cấu hình tỷ lệ độ khó và phân bổ môn học.</p>
              </header>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div className="card space-y-6">
                  <h3 className="font-bold text-slate-700 border-b pb-2">Thông tin chung</h3>
                  <div className="space-y-4">
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-slate-600 mb-1">Số câu Phần I</label>
                        <input 
                          type="number" 
                          className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                          value={structure.part1_count}
                          onChange={(e) => {
                            const val = parseInt(e.target.value) || 0;
                            setStructure(prev => ({ 
                              ...prev, 
                              part1_count: val,
                              total_questions: val + prev.part2_count + prev.part3_count 
                            }));
                          }}
                        />
                        <p className="text-[10px] text-slate-400 mt-1">Trắc nghiệm 1 lựa chọn</p>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-600 mb-1">Số câu Phần II</label>
                        <input 
                          type="number" 
                          className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                          value={structure.part2_count}
                          onChange={(e) => {
                            const val = parseInt(e.target.value) || 0;
                            setStructure(prev => ({ 
                              ...prev, 
                              part2_count: val,
                              total_questions: prev.part1_count + val + prev.part3_count 
                            }));
                          }}
                        />
                        <p className="text-[10px] text-slate-400 mt-1">Trắc nghiệm Đúng/Sai</p>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-600 mb-1">Số câu Phần III</label>
                        <input 
                          type="number" 
                          className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                          value={structure.part3_count}
                          onChange={(e) => {
                            const val = parseInt(e.target.value) || 0;
                            setStructure(prev => ({ 
                              ...prev, 
                              part3_count: val,
                              total_questions: prev.part1_count + prev.part2_count + val 
                            }));
                          }}
                        />
                        <p className="text-[10px] text-slate-400 mt-1">Câu hỏi Trả lời ngắn</p>
                      </div>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-slate-600 mb-1">Tổng cộng: <span className="font-bold text-primary">{structure.total_questions} câu</span></label>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-slate-600 mb-1">Danh sách mã đề</label>
                      <div className="flex flex-wrap gap-2 mb-2">
                        {examCodes.map((code, idx) => (
                          <div key={idx} className="relative group">
                            <input 
                              type="text" 
                              className="w-20 px-2 py-1 border border-slate-200 rounded text-center font-bold focus:ring-2 focus:ring-primary/20 outline-none"
                              value={code}
                              onChange={(e) => {
                                const newCodes = [...examCodes];
                                newCodes[idx] = e.target.value;
                                setExamCodes(newCodes);
                              }}
                            />
                            <button 
                              onClick={() => setExamCodes(examCodes.filter((_, i) => i !== idx))}
                              className="absolute -top-2 -right-2 bg-rose-500 text-white rounded-full p-0.5 opacity-0 group-hover:opacity-100 transition-opacity"
                            >
                              <Trash2 size={10} />
                            </button>
                          </div>
                        ))}
                        <button 
                          onClick={() => setExamCodes([...examCodes, (parseInt(examCodes[examCodes.length - 1] || '100') + 1).toString()])}
                          className="w-20 px-2 py-1 border border-dashed border-slate-300 rounded text-slate-400 hover:border-primary hover:text-primary transition-colors flex items-center justify-center"
                        >
                          <Plus size={16} />
                        </button>
                      </div>
                      <p className="text-[10px] text-slate-400 italic">Nhấp vào mã đề để sửa, di chuột để xóa.</p>
                    </div>
                  </div>
                </div>

                <div className="card space-y-6">
                  <h3 className="font-bold text-slate-700 border-b pb-2">Tỷ lệ độ khó (%)</h3>
                  <div className="space-y-4">
                    {Object.entries(structure.difficulty_ratios).map(([diff, val]) => (
                      <div key={diff}>
                        <div className="flex justify-between text-sm mb-1">
                          <span className="font-medium text-slate-600">{diff}</span>
                          <span className="font-bold text-primary">{val}%</span>
                        </div>
                        <input 
                          type="range" 
                          className="w-full h-2 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-primary"
                          value={val}
                          onChange={(e) => setStructure(prev => ({
                            ...prev,
                            difficulty_ratios: { ...prev.difficulty_ratios, [diff]: parseInt(e.target.value) }
                          }))}
                        />
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="card">
                <h3 className="font-bold text-slate-700 border-b pb-4 mb-4">Phân bổ môn học (Số câu)</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {Object.entries(structure.topic_distribution).map(([topic, count]) => (
                    <div key={topic} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border border-slate-100">
                      <span className="text-sm font-medium text-slate-600 truncate mr-2">{topic}</span>
                      <input 
                        type="number" 
                        className="w-16 px-2 py-1 border border-slate-200 rounded bg-white text-center text-sm font-bold"
                        value={count}
                        onChange={(e) => setStructure(prev => ({
                          ...prev,
                          topic_distribution: { ...prev.topic_distribution, [topic]: parseInt(e.target.value) }
                        }))}
                      />
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex justify-end">
                <button onClick={() => setActiveTab('mix')} className="btn-primary px-12 py-4 text-lg shadow-lg">
                  Lưu & Tiếp tục
                </button>
              </div>
            </motion.div>
          )}

          {activeTab === 'mix' && (
            <motion.div 
              key="mix"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="max-w-2xl mx-auto text-center py-12"
            >
              {!isMixing ? (
                <div className="space-y-8">
                  <div className="w-24 h-24 bg-blue-100 text-primary rounded-full flex items-center justify-center mx-auto mb-6">
                    <Shuffle size={48} />
                  </div>
                  <h2 className="text-3xl font-bold text-slate-800">Sẵn sàng trộn đề?</h2>
                  <p className="text-slate-500">
                    Hệ thống sẽ dựa trên cấu trúc đã thiết lập để chọn ngẫu nhiên {structure.total_questions} câu hỏi 
                    và tạo ra {examCodes.length} mã đề thi khác nhau.
                  </p>
                  
                  <div className="card text-left bg-blue-50 border-blue-100">
                    <h4 className="font-bold text-primary mb-2 flex items-center gap-2">
                      <CheckCircle2 size={18} /> Tóm tắt cấu hình
                    </h4>
                    <ul className="text-sm text-slate-600 space-y-1">
                      <li>• Tổng số câu: <strong>{structure.total_questions}</strong></li>
                      <li>• Số mã đề: <strong>{examCodes.length}</strong> ({examCodes.join(', ')})</li>
                      <li>• Môn học: <strong>{Object.keys(structure.topic_distribution).length}</strong></li>
                    </ul>
                  </div>

                  <div className="card text-left space-y-4">
                    <h4 className="font-bold text-slate-700 border-b pb-2">Tùy chỉnh tiêu đề đề thi</h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Tên đơn vị (Bên trái)</label>
                        <input 
                          type="text" 
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                          value={headerConfig.schoolName}
                          onChange={(e) => setHeaderConfig(prev => ({ ...prev, schoolName: e.target.value }))}
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Tên kỳ thi (Bên phải)</label>
                        <input 
                          type="text" 
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                          value={headerConfig.examName}
                          onChange={(e) => setHeaderConfig(prev => ({ ...prev, examName: e.target.value }))}
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Loại đề (Dưới tên đơn vị)</label>
                        <input 
                          type="text" 
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                          value={headerConfig.examType}
                          onChange={(e) => setHeaderConfig(prev => ({ ...prev, examType: e.target.value }))}
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Môn học (Dưới tên kỳ thi)</label>
                        <input 
                          type="text" 
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                          value={headerConfig.subject}
                          onChange={(e) => setHeaderConfig(prev => ({ ...prev, subject: e.target.value }))}
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Thời gian làm bài</label>
                        <input 
                          type="text" 
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                          value={headerConfig.time}
                          onChange={(e) => setHeaderConfig(prev => ({ ...prev, time: e.target.value }))}
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Số trang (Ghi chú)</label>
                        <input 
                          type="text" 
                          className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                          value={headerConfig.numPages}
                          onChange={(e) => setHeaderConfig(prev => ({ ...prev, numPages: e.target.value }))}
                        />
                      </div>

                      <h4 className="font-bold text-slate-700 border-b pb-2 pt-6">Tùy chỉnh tiêu đề các phần</h4>
                      <p className="text-[10px] text-slate-400 italic mb-4">Lưu ý: Tiêu đề sẽ tự động cập nhật khi bạn thay đổi số lượng câu hỏi ở bước trước. Bạn có thể sửa lại tại đây.</p>
                      <div className="space-y-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Tiêu đề Phần I</label>
                          <textarea 
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm h-20"
                            value={partLabels.part1}
                            onChange={(e) => setPartLabels(prev => ({ ...prev, part1: e.target.value }))}
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Tiêu đề Phần II</label>
                          <textarea 
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm h-20"
                            value={partLabels.part2}
                            onChange={(e) => setPartLabels(prev => ({ ...prev, part2: e.target.value }))}
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Tiêu đề Phần III</label>
                          <textarea 
                            className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm h-20"
                            value={partLabels.part3}
                            onChange={(e) => setPartLabels(prev => ({ ...prev, part3: e.target.value }))}
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  <button 
                    onClick={handleMix}
                    className="btn-primary w-full py-4 text-xl shadow-xl hover:scale-[1.02] active:scale-[0.98] transition-transform"
                  >
                    Bắt đầu trộn ngay
                  </button>
                </div>
              ) : (
                <div className="space-y-8">
                  <Loader2 className="w-24 h-24 text-primary animate-spin mx-auto mb-6" />
                  <h2 className="text-3xl font-bold text-slate-800">Đang trộn đề...</h2>
                  <div className="max-w-md mx-auto">
                    <div className="h-4 bg-slate-100 rounded-full overflow-hidden mb-2">
                      <motion.div 
                        className="h-full bg-primary"
                        initial={{ width: 0 }}
                        animate={{ width: `${mixProgress}%` }}
                      />
                    </div>
                    <p className="text-sm font-bold text-slate-400">{mixProgress}% Hoàn tất</p>
                  </div>
                  <p className="text-slate-500 italic">"Đang xáo trộn câu hỏi và đáp án..."</p>
                </div>
              )}
            </motion.div>
          )}

          {activeTab === 'ai' && (
            <motion.div 
              key="ai"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8"
            >
              <header>
                <h2 className="text-2xl font-bold text-slate-800">Trợ lý AI soạn đề</h2>
                <p className="text-slate-500">Sử dụng trí tuệ nhân tạo để tạo câu hỏi từ chủ đề hoặc nội dung văn bản.</p>
              </header>

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div className="lg:col-span-1 space-y-6">
                  <div className="card">
                    <h3 className="font-bold text-slate-700 mb-4 flex items-center gap-2">
                      <Wand2 size={18} className="text-primary" /> Yêu cầu AI
                    </h3>
                    <textarea 
                      className="w-full h-48 p-4 border border-slate-200 rounded-xl focus:ring-2 focus:ring-primary/20 outline-none resize-none text-sm"
                      placeholder="Ví dụ: Tạo 5 câu hỏi trắc nghiệm về lịch sử Việt Nam thời nhà Trần, độ khó trung bình..."
                      value={aiPrompt}
                      onChange={(e) => setAiPrompt(e.target.value)}
                    />
                    <button 
                      onClick={handleAiGenerate}
                      disabled={isAiGenerating || !aiPrompt.trim()}
                      className="btn-primary w-full mt-4 flex items-center justify-center gap-2"
                    >
                      {isAiGenerating ? (
                        <>
                          <Loader2 size={18} className="animate-spin" /> Đang tạo...
                        </>
                      ) : (
                        <>
                          <Sparkles size={18} /> Tạo câu hỏi ngay
                        </>
                      )}
                    </button>
                  </div>

                  <div className="bg-blue-50 p-4 rounded-xl border border-blue-100">
                    <h4 className="font-bold text-blue-800 text-sm mb-2 flex items-center gap-2">
                      <AlertCircle size={16} /> Mẹo soạn thảo
                    </h4>
                    <ul className="text-xs text-blue-700 space-y-2">
                      <li>• Cung cấp chủ đề cụ thể (Toán, Lý, Hóa...)</li>
                      <li>• Chỉ định số lượng câu hỏi mong muốn.</li>
                      <li>• Có thể dán một đoạn văn bản để AI tóm tắt thành câu hỏi.</li>
                      <li>• Yêu cầu rõ mức độ khó (Dễ, Khó).</li>
                    </ul>
                  </div>
                </div>

                <div className="lg:col-span-2">
                  <div className="card h-full flex flex-col min-h-[500px]">
                    <div className="flex justify-between items-center mb-6">
                      <h3 className="font-bold text-slate-700">Kết quả từ AI</h3>
                      {aiResult.length > 0 && (
                        <button 
                          onClick={importAiQuestions}
                          className="btn-primary flex items-center gap-2"
                        >
                          <Plus size={18} /> Thêm tất cả vào ngân hàng
                        </button>
                      )}
                    </div>

                    {aiResult.length > 0 ? (
                      <div className="space-y-4 overflow-y-auto max-h-[600px] pr-2">
                        {aiResult.map((q, idx) => (
                          <div key={idx} className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                            <div className="flex justify-between items-start mb-2">
                              <span className="text-xs font-bold px-2 py-1 bg-primary/10 text-primary rounded">Câu {idx + 1}</span>
                              <span className="text-xs font-medium text-slate-500">{q.difficulty} | {q.topic}</span>
                            </div>
                            <p className="font-medium text-slate-800 mb-3">{q.content}</p>
                            {q.image_url && (
                              <div className="mb-3 rounded-lg overflow-hidden border border-slate-200 bg-white">
                                <img 
                                  src={q.image_url} 
                                  alt="AI generated" 
                                  className="max-h-48 w-full object-contain"
                                  onError={(e) => (e.currentTarget.style.display = 'none')}
                                />
                              </div>
                            )}
                            <div className="grid grid-cols-2 gap-2 text-sm">
                              <div className={q.correct_answer === 'A' ? 'text-emerald-600 font-bold' : 'text-slate-600'}>A. {q.option_a}</div>
                              <div className={q.correct_answer === 'B' ? 'text-emerald-600 font-bold' : 'text-slate-600'}>B. {q.option_b}</div>
                              <div className={q.correct_answer === 'C' ? 'text-emerald-600 font-bold' : 'text-slate-600'}>C. {q.option_c}</div>
                              <div className={q.correct_answer === 'D' ? 'text-emerald-600 font-bold' : 'text-slate-600'}>D. {q.option_d}</div>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="flex-1 flex flex-col items-center justify-center text-slate-400">
                        <Sparkles size={48} className="mb-4 opacity-20" />
                        <p>Nhập yêu cầu và nhấn nút để bắt đầu sáng tạo</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'export' && (
            <motion.div 
              key="export"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8"
            >
              <header className="flex justify-between items-center">
                <div>
                  <h2 className="text-2xl font-bold text-slate-800">Kết quả trộn đề</h2>
                  <p className="text-slate-500">Đã tạo thành công {generatedExams.length} mã đề thi.</p>
                </div>
                <button onClick={exportAnswerKey} className="btn-secondary flex items-center gap-2 text-emerald-600 border-emerald-200 hover:bg-emerald-50">
                  <Download size={18} /> Xuất bảng đáp án (Excel)
                </button>
              </header>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {generatedExams.map((exam) => (
                  <div key={exam.code} className="card group hover:border-primary transition-colors">
                    <div className="flex justify-between items-start mb-4">
                      <div className="w-12 h-12 bg-slate-100 rounded-lg flex items-center justify-center font-bold text-slate-600 group-hover:bg-primary group-hover:text-white transition-colors">
                        {exam.code}
                      </div>
                      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Mã đề</span>
                    </div>
                    <h4 className="font-bold text-slate-800 mb-1">Đề thi mã {exam.code}</h4>
                    <div className="space-y-1 mb-6">
                      <p className="text-[10px] text-slate-500">Phần I: {exam.part1.length} câu (Trắc nghiệm 1 lựa chọn)</p>
                      <p className="text-[10px] text-slate-500">Phần II: {exam.part2.length} câu (Trắc nghiệm Đúng/Sai)</p>
                      <p className="text-[10px] text-slate-500">Phần III: {exam.part3.length} câu (Trả lời ngắn)</p>
                    </div>
                    
                    <div className="grid grid-cols-2 gap-2">
                      <button 
                        onClick={() => exportToWord(exam)}
                        className="btn-secondary py-2 text-xs flex items-center justify-center gap-1"
                      >
                        <FileText size={14} /> Word
                      </button>
                      <button 
                        onClick={() => exportToPDF(exam)}
                        className="btn-secondary py-2 text-xs flex items-center justify-center gap-1"
                      >
                        <Download size={14} /> PDF
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Modal */}
      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsModalOpen(false)}
              className="absolute inset-0 bg-slate-900/50 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden"
            >
                {modalType === 'delete' || modalType === 'bulk-delete' ? (
                  <div className="p-8 text-center">
                    <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-4">
                      <Trash2 size={32} />
                    </div>
                    <h3 className="text-xl font-bold text-slate-800 mb-2">Xác nhận xóa</h3>
                    <p className="text-slate-500 mb-8">
                      {modalType === 'delete' 
                        ? 'Bạn có chắc chắn muốn xóa câu hỏi này? Hành động này không thể hoàn tác.'
                        : `Bạn có chắc chắn muốn xóa ${selectedIds.length} câu hỏi đã chọn? Hành động này không thể hoàn tác.`
                      }
                    </p>
                    <div className="flex gap-4 justify-center">
                      <button onClick={() => setIsModalOpen(false)} className="btn-secondary px-8">Hủy</button>
                      <button 
                        onClick={modalType === 'delete' ? handleDelete : confirmBulkDelete} 
                        className="bg-rose-600 text-white px-8 py-2 rounded-lg font-medium hover:bg-rose-700 transition-colors"
                      >
                        Xóa ngay
                      </button>
                    </div>
                  </div>
                ) : (
                <div className="flex flex-col h-[80vh]">
                  <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                    <h3 className="text-xl font-bold text-slate-800">
                      {modalType === 'add' ? 'Thêm câu hỏi mới' : 'Chỉnh sửa câu hỏi'}
                    </h3>
                    <button onClick={() => setIsModalOpen(false)} className="text-slate-400 hover:text-slate-600">
                      <Plus size={24} className="rotate-45" />
                    </button>
                  </div>
                  <div className="flex-1 overflow-y-auto p-6 space-y-6">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-slate-600 mb-1">Loại câu hỏi</label>
                        <select 
                          className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                          value={currentQuestion.type}
                          onChange={(e) => {
                            const type = e.target.value as any;
                            setCurrentQuestion(prev => ({ 
                              ...prev, 
                              type,
                              correct_answer: type === 'TRUE_FALSE' ? 'T,T,T,T' : (type === 'SHORT_ANSWER' ? '' : 'A')
                            }));
                          }}
                        >
                          <option value="SINGLE_CHOICE">Trắc nghiệm 1 lựa chọn</option>
                          <option value="TRUE_FALSE">Trắc nghiệm Đúng/Sai</option>
                          <option value="SHORT_ANSWER">Trả lời ngắn</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-600 mb-1">Môn học</label>
                        <input 
                          type="text" 
                          className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                          value={currentQuestion.topic}
                          onChange={(e) => setCurrentQuestion(prev => ({ ...prev, topic: e.target.value }))}
                        />
                      </div>
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-slate-600 mb-1">Nội dung câu hỏi</label>
                      <ReactQuill 
                        {...({
                          ref: quillRef,
                          theme: "snow",
                          value: ensureMathAssetsRendered(currentQuestion.content || ''),
                          onChange: (content: string) => setCurrentQuestion(prev => ({ ...prev, content })),
                          modules: modules,
                          className: "h-64 mb-12"
                        } as any)}
                      />
                    </div>

                    {currentQuestion.type === 'SINGLE_CHOICE' && (
                      <div className="space-y-4">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div>
                            <label className="block text-sm font-medium text-slate-600 mb-1">Phương án A</label>
                            <input 
                              type="text" 
                              className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                              value={currentQuestion.option_a}
                              onChange={(e) => setCurrentQuestion(prev => ({ ...prev, option_a: e.target.value }))}
                            />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-600 mb-1">Phương án B</label>
                            <input 
                              type="text" 
                              className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                              value={currentQuestion.option_b}
                              onChange={(e) => setCurrentQuestion(prev => ({ ...prev, option_b: e.target.value }))}
                            />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-600 mb-1">Phương án C</label>
                            <input 
                              type="text" 
                              className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                              value={currentQuestion.option_c}
                              onChange={(e) => setCurrentQuestion(prev => ({ ...prev, option_c: e.target.value }))}
                            />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-600 mb-1">Phương án D</label>
                            <input 
                              type="text" 
                              className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                              value={currentQuestion.option_d}
                              onChange={(e) => setCurrentQuestion(prev => ({ ...prev, option_d: e.target.value }))}
                            />
                          </div>
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-slate-600 mb-1">Đáp án đúng</label>
                          <select 
                            className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                            value={currentQuestion.correct_answer}
                            onChange={(e) => setCurrentQuestion(prev => ({ ...prev, correct_answer: e.target.value }))}
                          >
                            <option value="A">A</option>
                            <option value="B">B</option>
                            <option value="C">C</option>
                            <option value="D">D</option>
                          </select>
                        </div>
                      </div>
                    )}

                    {currentQuestion.type === 'TRUE_FALSE' && (
                      <div className="space-y-4">
                        <p className="text-xs font-bold text-slate-400 uppercase">Các lệnh/phát biểu (a, b, c, d)</p>
                        {[0, 1, 2, 3].map((idx) => {
                          const labels = ['a', 'b', 'c', 'd'];
                          const field = `option_${labels[idx]}` as any;
                          const answers = (currentQuestion.correct_answer || 'T,T,T,T').split(',');
                          
                          return (
                            <div key={idx} className="flex gap-4 items-center">
                              <span className="font-bold text-slate-400 w-4">{labels[idx]})</span>
                              <input 
                                type="text" 
                                className="flex-1 px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                                value={(currentQuestion as any)[field]}
                                onChange={(e) => setCurrentQuestion(prev => ({ ...prev, [field]: e.target.value }))}
                              />
                              <select 
                                className="w-24 px-2 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                                value={answers[idx]}
                                onChange={(e) => {
                                  answers[idx] = e.target.value;
                                  setCurrentQuestion(prev => ({ ...prev, correct_answer: answers.join(',') }));
                                }}
                              >
                                <option value="T">Đúng</option>
                                <option value="F">Sai</option>
                              </select>
                            </div>
                          );
                        })}
                      </div>
                    )}

                    {currentQuestion.type === 'SHORT_ANSWER' && (
                      <div>
                        <label className="block text-sm font-medium text-slate-600 mb-1">Đáp án đúng (Trả lời ngắn)</label>
                        <input 
                          type="text" 
                          className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                          value={currentQuestion.correct_answer}
                          onChange={(e) => setCurrentQuestion(prev => ({ ...prev, correct_answer: e.target.value }))}
                          placeholder="Nhập đáp án chính xác..."
                        />
                      </div>
                    )}

                    <div>
                      <label className="block text-sm font-medium text-slate-600 mb-1">Độ khó</label>
                      <select 
                        className="w-full px-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-primary/20 outline-none"
                        value={currentQuestion.difficulty}
                        onChange={(e) => setCurrentQuestion(prev => ({ ...prev, difficulty: e.target.value as any }))}
                      >
                        <option value="Dễ">Dễ</option>
                        <option value="Trung bình">Trung bình</option>
                        <option value="Khó">Khó</option>
                      </select>
                    </div>
                  </div>
                  <div className="p-6 border-t border-slate-100 flex gap-4 justify-end">
                    <button onClick={() => setIsModalOpen(false)} className="btn-secondary px-8">Hủy</button>
                    <button onClick={handleSave} className="btn-primary px-8">Lưu thay đổi</button>
                  </div>
                </div>
              )}
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
