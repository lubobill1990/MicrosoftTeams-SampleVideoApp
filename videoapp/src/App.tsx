import { useCallback, useEffect, useRef, useState } from 'react'
import './App.css'
import { app, video } from "@microsoft/teams-js";

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
    app.initialize([
      "https://microsoft.github.io"
    ]).then(() => {
      initializeVideoApp()
    });
  }, [initializeVideoApp]);

  return <>
    <div className="card">
      <button onClick={() => setCount((count) => count + 1)}>
        count is {count}
      </button>
    </div>
  </>
}

export default App
