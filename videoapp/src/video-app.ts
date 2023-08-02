import { video } from "@microsoft/teams-js";
import { WebGL2Grayscale } from "./webgl2";

const ZERO_DELAY_EMPTY_EFFECT_ID = "00000000-0000-0000-0000-000000000000";
const GRAYSCALE_EFFECT_ID = "00000000-0000-0000-0001-000000000000";
const EFFECT_ID_FAILED_TO_LOAD_ASSET = "00000000-0000-0001-0000-000000000000";

export class VideoApp {
  private grayscaleProcessor: WebGL2Grayscale | null = null;
  private selectedEffectId: string | undefined = undefined;

  videoEffectSelected(effectId: string | undefined) {
    if (effectId === EFFECT_ID_FAILED_TO_LOAD_ASSET) {
      // When this effect require some assets to be loaded, but failed to load them,
      // we should reject the promise, so that Teams can handle the error properly.
      console.log(
        `Failed to prepare required resources for ${effectId}. Rejecting the promise.`
      );
      return Promise.reject();
    }
    this.selectedEffectId = effectId;
    if (effectId === GRAYSCALE_EFFECT_ID) {
      if (this.grayscaleProcessor === null) {
        this.grayscaleProcessor = new WebGL2Grayscale();
        this.grayscaleProcessor.setup();
      }
    } else if (this.grayscaleProcessor !== null) {
      this.grayscaleProcessor.tearDown();
      this.grayscaleProcessor = null;
    }
    console.log(
      `New effect selected ${effectId}. We are ready to process video frames using this effect. Resolving the promise.`
    );
    return Promise.resolve();
  }

  private getLatencyFromEffectId(effectId: string): number {
    if (effectId.indexOf("00000000-0000-0000-0000-") === 0) {
      return parseInt(effectId.split("-")[4]);
    }
    return 0;
  }

  async videoFrameHandler(frame: video.VideoFrameData): Promise<VideoFrame> {
    const effectId = this.selectedEffectId;
    const videoFrame = frame.videoFrame as VideoFrame;
    if (!effectId || effectId === ZERO_DELAY_EMPTY_EFFECT_ID) {
      return videoFrame;
    }
    const latency = this.getLatencyFromEffectId(effectId);

    if (latency > 0) {
      return new Promise((resolve) => {
        setTimeout(() => {
          resolve(videoFrame);
        }, latency);
      });
    }

    if (effectId === GRAYSCALE_EFFECT_ID) {
      if (this.grayscaleProcessor === null) {
        throw "Grayscale effect is not initialized";
      }

      this.grayscaleProcessor.draw(videoFrame);
      const grayscaledVideoFrame = new VideoFrame(
        this.grayscaleProcessor.getCanvas(),
        {
          displayHeight: videoFrame.displayHeight, // Should keep the same size as the original video frame
          displayWidth: videoFrame.displayWidth,
          timestamp: videoFrame.timestamp,
        }
      );

      return grayscaledVideoFrame;
    }
    // Throw error if the effect id is not recognized or some other error happens, so that Teams can handle the error properly.
    throw `Unknown effect id ${effectId}`;
  }

  async videoBufferHandler(
    videoBufferData: video.VideoBufferData,
    notifyVideoFrameProcessed: () => void,
    _notifyError: (error: string) => void
  ) {
    const effectId = this.selectedEffectId;
    if (!effectId || effectId === ZERO_DELAY_EMPTY_EFFECT_ID) {
      notifyVideoFrameProcessed();
      return;
    }

    const latency = this.getLatencyFromEffectId(effectId);
    if (latency > 0) {
      setTimeout(() => {
        notifyVideoFrameProcessed();
      }, latency);
      return;
    }

    const frame = new VideoFrame(videoBufferData.videoFrameBuffer, {
      format: "NV12",
      codedWidth: videoBufferData.width,
      codedHeight: videoBufferData.height,
      timestamp: Date.now(),
    });
    const outputFrame = await this.videoFrameHandler({
      videoFrame: frame,
    });
    const bitmap = await createImageBitmap(outputFrame);
    console.log(bitmap.width);
    //TODO convert bitmap pixels to NV12 and copy back to videoBufferData.videoFrameBuffer
    notifyVideoFrameProcessed();
    frame.close();
    outputFrame.close();
  }
}
