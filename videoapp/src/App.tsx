import { useCallback, useEffect, useRef, useState } from 'react'
import './App.css'
import { app, video } from "@microsoft/teams-js";
import { WebGL2Grayscale } from './webgl2';

function App() {
  const [count, setCount] = useState(0)
  const effectIdRef = useRef<string | undefined>(undefined);
  const initializeVideoApp = useCallback(() => {
    video.registerForVideoEffect((effectId: string | undefined) => {
      console.log(`New effect selected ${effectId}`);
      effectIdRef.current = effectId;
      return Promise.resolve();
    });
    video.mediaStream.registerForVideoFrame((frame: any) => {
      const effectId = effectIdRef.current;
      const videoFrame = frame.videoFrame as VideoFrame;
      if (effectId?.indexOf("00000000-0000-0000-0000-0000") === 0) {
        const latency = parseInt(effectId.split('-')[4]);
        return new Promise((resolve) => {
          setTimeout(() => {
            resolve(videoFrame);
          }, latency);
        });
      }
      if (effectId === '00000000-0000-0000-0001-000000000000') {

      }
      return Promise.resolve(videoFrame);
    })
  }, [effectIdRef])


  useEffect(() => {
    return;
    app.initialize([
      "https://microsoft.github.io"
    ]).then(() => {
      initializeVideoApp()
    });
  }, [initializeVideoApp]);

  useEffect(() => {
    if (count === 0) {
      return;
    }
    navigator.mediaDevices.getDisplayMedia({}).then((stream) => {
      // turn to VideoFrame using TransformStream
      const videoTrack = stream.getVideoTracks()[0];
      // initialize a MediaStreamTrackProcessor from the videoTrack
      const processor = new MediaStreamTrackProcessor({ track: videoTrack });
      // initialize a MediaStreamTrackGenerator
      const generator = new MediaStreamTrackGenerator({ kind: 'video' });

      const grayscale = new WebGL2Grayscale();
      grayscale.setup();
      const transformStream = new TransformStream({
        transform: async (chunk: VideoFrame, controller) => {
          const imageBitmap = await createImageBitmap(chunk);
          grayscale.draw(imageBitmap);
          const grayscaled = new VideoFrame(grayscale.getCanvas(), {
            timestamp: chunk.timestamp,
            displayHeight: chunk.displayHeight,
            displayWidth: chunk.displayWidth,
          });
          chunk.close();
          controller.enqueue(grayscaled);
        }
      });
      const readableStream = processor.readable;
      const writableStream = generator.writable;
      readableStream.pipeThrough(transformStream).pipeTo(writableStream);
      const outputMediaStream = new MediaStream([generator]);
      videoRef.current!.srcObject = outputMediaStream;
    })
  }, [count === 0]);

  const videoRef = useRef<HTMLVideoElement>(null);
  return <>
    <div className="card">
      <button onClick={() => setCount((count) => count + 1)}>
        count is {count}
      </button>
      <video autoPlay ref={videoRef} style={{ width: "100%" }}></video>
    </div>
  </>
}

export default App
