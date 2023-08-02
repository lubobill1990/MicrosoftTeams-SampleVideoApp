export class PerformanceMonitor {
  private queue: number[] = [];
  private processingFrameTimestamp = 0;
  private startProcessingTime = 0;
  private lastPrintTime = 0;
  startFrame(videoFrame: VideoFrame) {
    this.startProcessingTime = performance.now();
    this.processingFrameTimestamp = videoFrame.timestamp;
  }

  endFrame(videoFrame: VideoFrame) {
    if (videoFrame.timestamp !== this.processingFrameTimestamp) {
      return;
    }
    this.queue.push(performance.now() - this.startProcessingTime);
    if (this.queue.length > 100) {
      this.queue.splice(0, this.queue.length - 100);
    }
    if (performance.now() - this.lastPrintTime > 1000) {
      this.lastPrintTime = performance.now();
      console.log(
        `Average processing time: ${(this.queue.reduce((a, b) => a + b, 0) /
          this.queue.length).toFixed(2)}ms for recent ${this.queue.length} frames`
      );
    }
  }

  reset() {
    this.queue = [];
    this.processingFrameTimestamp = 0;
    this.startProcessingTime = 0;
    this.lastPrintTime = 0;
  }
}
