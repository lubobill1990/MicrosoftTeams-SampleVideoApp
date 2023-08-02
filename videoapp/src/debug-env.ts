import { PerformanceMonitor } from "./performance-monitor";
import { VideoApp } from "./video-app";

export class DebugEnv {
  private debugEnabled = false;
  private debugVideoPipelineRunning = false;
  private performanceMonitor = new PerformanceMonitor();
  start() {
    document.getElementById("enableDebug")?.addEventListener("click", (e) => {
      this.debugEnabled = (e.target as HTMLInputElement).checked;
      this.updateVideoWrap();
    });

    document
      .getElementById("startLocalCamera")!
      .addEventListener("click", async () => {
        if (!this.debugEnabled || this.debugVideoPipelineRunning) {
          return;
        }
        this.debugVideoPipelineRunning = true;
        this.startVideoPipeline();
      });
  }

  updateVideoWrap() {
    const videoWrap = document.getElementById("videoWrap")!;

    if (this.debugEnabled) {
      videoWrap.style.setProperty("display", "block");
    } else {
      videoWrap.style.setProperty("display", "none");
    }
  }

  async startVideoPipeline() {
    const videoApp = new VideoApp();

    const inputMediaStream = await navigator.mediaDevices.getUserMedia({
      video: true,
    });
    const processedVideoElement = document.getElementById(
      "processedVideo"
    ) as HTMLVideoElement;
    const originVideoElement = document.getElementById(
      "originVideo"
    ) as HTMLVideoElement;
    const track = inputMediaStream.getVideoTracks()[0];
    const processor = new MediaStreamTrackProcessor({ track });
    const generator = new MediaStreamTrackGenerator({ kind: "video" });
    const outputMediaStream = new MediaStream([generator]);
    processor.readable
      .pipeThrough(
        new TransformStream({
          transform: async (videoFrame, controller) => {
            try {
              this.performanceMonitor.startFrame(videoFrame);
              const processedFrame = await videoApp.videoFrameHandler({
                videoFrame,
              });
              this.performanceMonitor.endFrame(videoFrame);
              if (processedFrame !== videoFrame) {
                videoFrame.close();
              }
              controller.enqueue(processedFrame);
            } catch (e) {
              controller.terminate();
              alert("Failed to process video frame. Terminate video pipeline.");
            }
          },
        })
      )
      .pipeTo(generator.writable);
    originVideoElement.srcObject = inputMediaStream;
    processedVideoElement.srcObject = outputMediaStream;

    document.addEventListener("click", (e) => {
      if ((e.target as Element)?.tagName !== "BUTTON") {
        return;
      }
      const clickedButton = e.target as HTMLButtonElement;
      // Effect picked by user in video app's UI can't be applied immediately.
      // Should send effect change notification to Teams,
      // and wait for Teams to call `vdieoApp.videoEffectSelected()` callback registered through `registerForVideoEffect`.
      videoApp.videoEffectSelected(clickedButton.dataset.id);
      this.performanceMonitor.reset();
    });
  }
}
