import { useCallback, useEffect, useRef } from 'react'
import './App.css'
import { app, video } from "@microsoft/teams-js";
import { WebGL2Grayscale } from './webgl2';

function App() {
  const effectIdRef = useRef<string | undefined>(undefined);
  const initializeVideoApp = useCallback(() => {
    let grayscale: WebGL2Grayscale | null = null;
    video.registerForVideoEffect((effectId: string | undefined) => {
      console.log(`New effect selected ${effectId}`);
      effectIdRef.current = effectId;
      if (effectId === '00000000-0000-0000-0001-000000000000') {
        if (grayscale === null) {
          grayscale = new WebGL2Grayscale();
          grayscale.setup();
        }
      } else {
        grayscale = null;
      }
      return Promise.resolve();
    });
    video.mediaStream.registerForVideoFrame(async (frame: any) => {
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
        if (grayscale === null) {
          throw 'Grayscale effect is not initialized';
        }
        const imageBitmap = await createImageBitmap(videoFrame);
        grayscale.draw(imageBitmap);
        const grayscaled = new VideoFrame(grayscale.getCanvas(), {
          timestamp: videoFrame.timestamp,
          displayHeight: videoFrame.displayHeight,
          displayWidth: videoFrame.displayWidth,
        });
        videoFrame.close();
        return grayscaled;
      }
      return Promise.resolve(videoFrame);
    })
  }, [effectIdRef])

  useEffect(() => {
    app.initialize([
      "https://microsoft.github.io"
    ]).then(() => {
      initializeVideoApp()
    });
  }, [initializeVideoApp]);

  return <>
    <div className="card">
      <h1>This is a Microsoft Teams sample video app</h1>
    </div>
  </>
}

export default App
