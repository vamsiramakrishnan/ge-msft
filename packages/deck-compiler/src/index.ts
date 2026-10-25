import PptxGenJS from 'pptxgenjs';
import { DeckSpecSchema, type DeckElement, type DeckSpec, type DeckTextStyle } from '@ge/contracts';

export interface CompiledDeckArtifact {
  readonly base64: string;
  readonly format: 'pptx';
  readonly slideCount: number;
  readonly specFingerprint: string;
  readonly warnings: readonly string[];
}

export interface CompileDeckOptions {
  readonly compression?: boolean;
  /** Fail before generating a file when geometry overlaps or extends outside the slide. */
  readonly strictLayout?: boolean;
}

interface SlideSize {
  readonly width: number;
  readonly height: number;
}

interface Frame {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

const WIDE_SIZE: SlideSize = { width: 13.333, height: 7.5 };
const STANDARD_SIZE: SlideSize = { width: 10, height: 7.5 };

export async function compileDeckSpecToBase64(
  input: unknown,
  options: CompileDeckOptions = {},
): Promise<CompiledDeckArtifact> {
  const spec = DeckSpecSchema.parse(input);
  const warnings: string[] = [];
  if (options.strictLayout) {
    const issues = inspectDeckLayout(spec);
    if (issues.length) throw new Error(`Deck layout failed: ${issues.join('; ')}`);
  }
  const pptx = new PptxGenJS();
  pptx.layout = spec.layout === 'standard' ? 'LAYOUT_4x3' : 'LAYOUT_WIDE';
  pptx.author = spec.author ?? 'Gemini Enterprise M365 Add-in';
  if (spec.subject) pptx.subject = spec.subject;
  if (spec.title) pptx.title = spec.title;
  pptx.company = 'Google Gemini Enterprise for Microsoft 365';
  pptx.theme = {
    headFontFace: spec.theme?.headFontFace ?? 'Aptos Display',
    bodyFontFace: spec.theme?.bodyFontFace ?? 'Aptos',
  };

  const size = spec.layout === 'standard' ? STANDARD_SIZE : WIDE_SIZE;
  for (const slideSpec of spec.slides) {
    const slide = pptx.addSlide();
    if (slideSpec.backgroundColor) slide.background = { color: slideSpec.backgroundColor };
    if (slideSpec.title) {
      slide.addText(slideSpec.title, {
        x: 0.55,
        y: 0.35,
        w: size.width - 1.1,
        h: 0.52,
        fontFace: spec.theme?.headFontFace ?? 'Aptos Display',
        fontSize: 24,
        bold: true,
        color: spec.theme?.headingColor ?? '1F1F1F',
        margin: 0,
        fit: 'shrink',
      });
    }
    if (slideSpec.subtitle) {
      slide.addText(slideSpec.subtitle, {
        x: 0.58,
        y: 0.92,
        w: size.width - 1.16,
        h: 0.4,
        fontFace: spec.theme?.bodyFontFace ?? 'Aptos',
        fontSize: 11,
        color: '595959',
        margin: 0,
        fit: 'shrink',
      });
    }
    if (slideSpec.notes) slide.addNotes(slideSpec.notes);
    for (const element of slideSpec.elements) {
      addElement(slide, element, size, warnings, spec.theme);
    }
  }

  const output = await pptx.write({
    outputType: 'base64',
    compression: options.compression ?? true,
  });
  if (typeof output !== 'string') {
    throw new Error('PptxGenJS did not return a base64 string for outputType=base64');
  }

  return {
    base64: output,
    format: 'pptx',
    slideCount: spec.slides.length,
    specFingerprint: fingerprintDeckSpec(spec),
    warnings,
  };
}

export function fingerprintDeckSpec(spec: DeckSpec): string {
  return fnv1a32(stableStringify(DeckSpecSchema.parse(spec)));
}
/** Deterministic geometry/accessibility checks. Text overflow still requires native/render acceptance. */
export function inspectDeckLayout(input: unknown): string[] {
  const spec = DeckSpecSchema.parse(input);
  const size = spec.layout === 'standard' ? STANDARD_SIZE : WIDE_SIZE;
  const issues: string[] = [];
  for (const [slideIndex, slide] of spec.slides.entries()) {
    const boxes: Array<Frame & { name: string; allowOverlap?: boolean }> = [];
    if (slide.title) boxes.push({ name: 'title', x: 0.55, y: 0.35, w: size.width - 1.1, h: 0.52 });
    if (slide.subtitle)
      boxes.push({ name: 'subtitle', x: 0.58, y: 0.92, w: size.width - 1.16, h: 0.4 });
    for (const [index, element] of slide.elements.entries()) {
      const label = `Slide ${slideIndex + 1} element ${index + 1}`;
      if (
        element.w <= 0 ||
        element.h <= 0 ||
        element.x + element.w > size.width ||
        element.y + element.h > size.height
      )
        issues.push(`${label} extends outside the slide or has no area`);
      if (element.kind === 'image' && !element.altText?.trim())
        issues.push(`${label} needs alternative text`);
      boxes.push({ ...element, name: `element ${index + 1}` });
    }
    for (let left = 0; left < boxes.length; left++)
      for (let right = left + 1; right < boxes.length; right++) {
        const a = boxes[left]!;
        const b = boxes[right]!;
        if (a.allowOverlap || b.allowOverlap) continue;
        if (a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h)
          issues.push(`Slide ${slideIndex + 1}: ${a.name} overlaps ${b.name}`);
      }
  }
  return issues;
}

function addElement(
  slide: PptxGenJS.Slide,
  element: DeckElement,
  size: SlideSize,
  warnings: string[],
  theme?: DeckSpec['theme'],
): void {
  const frame = clampFrame(element, size, warnings);
  switch (element.kind) {
    case 'text':
      slide.addText(element.text, {
        ...frame,
        ...textOptions(element.style, theme),
        margin: 0.04,
        fit: 'shrink',
      });
      return;
    case 'bullets':
      slide.addText(
        element.items.map((text) => ({ text, options: { bullet: { type: 'bullet' } } })),
        {
          ...frame,
          ...textOptions(element.style, theme),
          breakLine: false,
          fit: 'shrink',
        },
      );
      return;
    case 'table':
      slide.addTable(
        element.rows.map((row) => row.map((text) => ({ text }))),
        {
          ...frame,
          ...textOptions(element.style, theme),
          border: { type: 'solid', color: 'D0D7DE', pt: 0.5 },
          fill: { color: 'FFFFFF' },
          margin: 0.05,
        },
      );
      return;
    case 'chart':
      slide.addChart(
        element.chartType === 'column' ? 'bar' : element.chartType,
        element.series.map((series) => ({
          name: series.name,
          labels: element.labels,
          values: series.values,
        })),
        {
          ...frame,
          barDir: element.chartType === 'bar' ? 'bar' : 'col',
          showLegend: element.showLegend,
          showTitle: Boolean(element.title),
          title: element.title,
          chartColors: theme?.accentColors,
          showValue: false,
          catAxisLabelFontFace: theme?.bodyFontFace ?? 'Aptos',
          valAxisLabelFontFace: theme?.bodyFontFace ?? 'Aptos',
        },
      );
      return;
    case 'image':
      slide.addImage({
        ...frame,
        data: `data:${element.mimeType};base64,${element.base64}`,
        altText: element.altText,
      });
      return;
  }
}

function textOptions(
  style: DeckTextStyle | undefined,
  theme?: DeckSpec['theme'],
): PptxGenJS.TextPropsOptions {
  return {
    fontFace: style?.fontFace ?? theme?.bodyFontFace ?? 'Aptos',
    fontSize: style?.fontSize ?? 14,
    bold: style?.bold,
    italic: style?.italic,
    color: style?.color ?? theme?.bodyColor ?? '1F1F1F',
    align: style?.align,
  };
}

function clampFrame(input: Frame, size: SlideSize, warnings: string[]): Frame {
  const x = clamp(input.x, 0, size.width - 0.05);
  const y = clamp(input.y, 0, size.height - 0.05);
  const w = clamp(input.w, 0.05, size.width - x);
  const h = clamp(input.h, 0.05, size.height - y);
  if (x !== input.x || y !== input.y || w !== input.w || h !== input.h) {
    warnings.push(`Adjusted an element frame to fit within ${size.width}x${size.height} slides.`);
  }
  return { x, y, w, h };
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function stableStringify(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map((item) => stableStringify(item)).join(',')}]`;
  if (value !== null && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    return `{${Object.keys(record)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

function fnv1a32(input: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}
