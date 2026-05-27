// src/lib/slideBuilder.ts
import pptxgen from "pptxgenjs";
import { addLogoToSlide } from "./helpers";

export class SlideBuilder {
  private slide: any;
  private pptx: pptxgen;

  constructor(pptx: pptxgen) {
    this.pptx = pptx;
    this.slide = pptx.addSlide();
  }

  /** Set slide background color */
  setBackground(color: string) {
    this.slide.background = { color };
    return this;
  }

  /** Add Ericsson logo */
  addLogo(isDarkBg: boolean) {
    addLogoToSlide(this.slide, isDarkBg);
    return this;
  }

  /** Add text with options */
  addText(text: string, opts: any) {
    this.slide.addText(text, opts);
    return this;
  }

  /** Add image */
  addImage(opts: any) {
    this.slide.addImage(opts);
    return this;
  }

  /** Add shape */
  addShape(shapeType: any, opts: any) {
    this.slide.addShape(shapeType, opts);
    return this;
  }

  /** Build and return the slide */
  build() {
    return this.slide;
  }
}
