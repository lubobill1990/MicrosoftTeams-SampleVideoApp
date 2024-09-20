import html2canvas from 'html2canvas';
import { App, createApp } from 'vue';
import CountDown from './vue/countdown.vue';

export class Canvas2D {
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private app: App<Element> | null = null;

  constructor() {
    const canvas = document.createElement('canvas');
    // create a WebGL2 context
    const ctx = canvas.getContext('2d');

    if (!ctx) {
      throw new Error('Canvas 2d context not supported');
    }

    this.canvas = canvas;
    this.ctx = ctx;
  }

  public setup(countDown: number) {
    this.app = createApp(CountDown, {
      countDown,
    });
    this.app.mount('#vue-app');
  }

  public tearDown() {
    this.app?.unmount();
  }

  async draw(videoFrame: VideoFrame | ImageBitmap) {
    const { ctx } = this;
    let width =
      videoFrame instanceof VideoFrame
        ? videoFrame.displayWidth
        : videoFrame.width;
    let height =
      videoFrame instanceof VideoFrame
        ? videoFrame.displayHeight
        : videoFrame.height;
    this.canvas.width = width;
    this.canvas.height = height;

    ctx.drawImage(videoFrame, 0, 0, ctx.canvas.width, ctx.canvas.height);
    const htmlDom = document.getElementById('html2canvas');
    if (htmlDom) {
      const htmlcanvas = await html2canvas(htmlDom, {
        backgroundColor: null,
      });
      // get the height width of the htmlDom
      const htmlDomWidth = htmlDom.clientWidth;
      const htmlDomHeight = htmlDom.clientHeight;
      const domInCanvasWidth = ctx.canvas.width / 2;
      const domInCanvasHeight =
        (domInCanvasWidth * htmlDomHeight) / htmlDomWidth;
      ctx.drawImage(
        htmlcanvas,
        (ctx.canvas.width - domInCanvasWidth) / 2,
        (ctx.canvas.height - domInCanvasHeight) / 2,
        domInCanvasWidth,
        domInCanvasHeight
      );
    }
  }

  getCanvas() {
    return this.canvas;
  }
}
