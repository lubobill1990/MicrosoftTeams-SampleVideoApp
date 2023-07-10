/* eslint-disable @typescript-eslint/no-unused-vars */
import { useCallback, useEffect, useRef } from "react";
import "./App.css";
import { app, video } from "@microsoft/teams-js";
import { WebGL2Grayscale } from "./webgl2";

function getLatencyFromEffectId(effectId: string): number {
  if (effectId.indexOf("00000000-0000-0000-0000-0000") === 0) {
    return parseInt(effectId.split("-")[4]);
  }
  return 0;
}

function App() {
  const canvasWrapRef = useRef<HTMLDivElement>(null);
  const originCanvasRef = useRef<HTMLCanvasElement>(null);
  const effectIdRef = useRef<string | undefined>(undefined);
  const initializeVideoApp = useCallback(() => {
    let grayscale: WebGL2Grayscale | null = null;
    let currentSize = "";
    let ctx2d = originCanvasRef.current?.getContext("2d");
    video.registerForVideoEffect((effectId: string | undefined) => {
      console.log(`New effect selected ${effectId}`);
      if (effectId === "00000000-0000-0001-0000-000000000000") {
        return Promise.reject();
      }
      effectIdRef.current = effectId;
      if (effectId === "00000000-0000-0000-0001-000000000000") {
        if (grayscale === null) {
          grayscale = new WebGL2Grayscale();
          grayscale.setup();
          canvasWrapRef.current?.appendChild(grayscale.getCanvas());
          console.log("Append canvas");
        }
      } else {
        grayscale = null;
      }
      return Promise.resolve();
    });

    const videoFrameHandler = async (
      frame: video.VideoFrameData
    ): Promise<VideoFrame> => {
      const effectId = effectIdRef.current;
      const videoFrame = frame.videoFrame as VideoFrame;
      if (!effectId || effectId === "00000000-0000-0000-0000-000000000000") {
        return videoFrame;
      }
      if (
        currentSize !== `${videoFrame.displayWidth}x${videoFrame.displayHeight}`
      ) {
        const newSize = `${videoFrame.displayWidth}x${videoFrame.displayHeight}`;
        console.log(`Input size changed from ${currentSize} to ${newSize}`);
        currentSize = newSize;
      }

      const latency = getLatencyFromEffectId(effectId);
      if (latency > 0) {
        return new Promise((resolve) => {
          setTimeout(() => {
            resolve(videoFrame);
          }, latency);
        });
      }
      
      if (effectId === "00000000-0000-0000-0001-000000000000") {
        if (grayscale === null) {
          throw "Grayscale effect is not initialized";
        }

        const imageBitmap = await createImageBitmap(videoFrame);
        // Draw original video frame
        if (
          originCanvasRef.current &&
          videoFrame.displayWidth !== originCanvasRef.current.width
        ) {
          originCanvasRef.current.width = videoFrame.displayWidth;
          originCanvasRef.current.height = videoFrame.displayHeight;
          ctx2d = originCanvasRef.current?.getContext("2d");
        }
        ctx2d?.drawImage(
          imageBitmap,
          0,
          0,
          videoFrame.displayWidth,
          videoFrame.displayHeight
        );
        grayscale.draw(imageBitmap);
        const grayscaled = new VideoFrame(grayscale.getCanvas(), {
          timestamp: videoFrame.timestamp,
          displayHeight: videoFrame.displayHeight,
          displayWidth: videoFrame.displayWidth,
        });
        videoFrame.close();
        return grayscaled;
      }
      return videoFrame;
    };

    video.registerForVideoFrame({
      videoFrameHandler,
      /**
       * Callback function to process the video frames shared by the host.
       */
      videoBufferHandler: async (
        videoBufferData: video.VideoBufferData,
        notifyVideoFrameProcessed,
        _notifyError
      ) => {
        const effectId = effectIdRef.current;
        if (!effectId || effectId === "00000000-0000-0000-0000-000000000000") {
           notifyVideoFrameProcessed();
           return;
        }

        const latency = getLatencyFromEffectId(effectId);
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
        const outputFrame = await videoFrameHandler({
          videoFrame: frame,
        });
        const bitmap = await createImageBitmap(outputFrame);
        console.log(bitmap.width);
        //TODO convert bitmap pixels to NV12 and copy back to videoBufferData.videoFrameBuffer
        notifyVideoFrameProcessed();
        frame.close();
        outputFrame.close();
      },
      /**
       * Video frame configuration supplied to the host to customize the generated video frame parameters, like format
       */
      config: {
        format: video.VideoFrameFormat.NV12,
      },
    });
  }, [effectIdRef]);

  useEffect(() => {
    app.initialize(["https://microsoft.github.io"]).then(() => {
      initializeVideoApp();
    });
  }, [initializeVideoApp]);

  return (
    <>
      <div className="card">
        <h1>This is a Microsoft Teams sample video app</h1>
        <h3>Original video</h3>
        <canvas ref={originCanvasRef} style={{ width: "100%" }}></canvas>
        <h3>Processed video</h3>
        <div ref={canvasWrapRef} style={{ width: "100%" }}></div>
      </div>
    </>
  );
}

export default App;
