import { useCallback, useEffect, useRef } from 'react'
import './App.css'
import { app, video } from "@microsoft/teams-js";
import { WebGL2Grayscale } from './webgl2';

function App() {
  const canvasWrapRef = useRef<HTMLDivElement>(null);
  const originCanvasRef = useRef<HTMLCanvasElement>(null);
  const effectIdRef = useRef<string | undefined>(undefined);
  const initializeVideoApp = useCallback(() => {
    let grayscale: WebGL2Grayscale | null = null;
    let currentSize = '';
    let ctx2d = originCanvasRef.current?.getContext('2d');
    video.registerForVideoEffect((effectId: string | undefined) => {
      console.log(`New effect selected ${effectId}`);
      effectIdRef.current = effectId;
      if (effectId === '00000000-0000-0000-0001-000000000000') {
        if (grayscale === null) {
          grayscale = new WebGL2Grayscale();
          grayscale.setup();
          canvasWrapRef.current?.appendChild(grayscale.getCanvas());
          console.log('Append canvas');
        }
      } else {
        grayscale = null;
      }
      return Promise.resolve();
    });
    video.mediaStream.registerForVideoFrame(async (frame: any) => {
      const effectId = effectIdRef.current;
      const videoFrame = frame.videoFrame as VideoFrame;
      if (currentSize !== `${videoFrame.displayWidth}x${videoFrame.displayHeight}`) {
        let newSize = `${videoFrame.displayWidth}x${videoFrame.displayHeight}`;
        console.log(`Input size changed from ${currentSize} to ${newSize}`);
        currentSize = newSize;
      }
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
        if (originCanvasRef.current && videoFrame.displayWidth !== originCanvasRef.current.width) {
          originCanvasRef.current.width = videoFrame.displayWidth;
          originCanvasRef.current.height = videoFrame.displayHeight;
          ctx2d = originCanvasRef.current?.getContext('2d');
        }
        ctx2d?.drawImage(imageBitmap, 0, 0, videoFrame.displayWidth, videoFrame.displayHeight);
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
      <canvas ref={originCanvasRef} style={{ width: '100%' }}></canvas>
      <div ref={canvasWrapRef} style={{ width: '100%' }}></div>
    </div>
  </>
}

export default App
