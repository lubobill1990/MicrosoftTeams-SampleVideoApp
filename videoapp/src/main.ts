import "./style.css";
import { app, video } from "@microsoft/teams-js";
import { VideoApp } from "./video-app";
import { DebugEnv } from "./debug-env";


function initializeVideoApp() {
  const videoApp = new VideoApp(localStorage.getItem('enableTimestampLog') === 'true');

  document.addEventListener("click", (e) => {
    if ((e.target as Element)?.tagName !== "BUTTON") {
      return;
    }
    const clickedButton = e.target as HTMLButtonElement;
    // Effect picked by user in video app's UI can't be applied immediately.
    // Should send effect change notification to Teams,
    // and wait for Teams to call `vdieoApp.videoEffectSelected()` callback registered through `registerForVideoEffect`.
    video.notifySelectedVideoEffectChanged(
      video.EffectChangeType.EffectChanged,
      clickedButton.dataset.id
    );
  });

  video.registerForVideoEffect(videoApp.videoEffectSelected.bind(videoApp));

  video.registerForVideoFrame({
    videoFrameHandler: videoApp.videoFrameHandler.bind(videoApp),
    /**
     * Callback function to process the video frames shared by the host.
     */
    videoBufferHandler: videoApp.videoBufferHandler.bind(videoApp),
    /**
     * Video frame configuration supplied to the host to customize the generated video frame parameters, like format
     */
    config: {
      format: video.VideoFrameFormat.NV12,
    },
  });
}

app.initialize(["https://microsoft.github.io"]).then(() => {
  initializeVideoApp();
}).catch(() => {
  console.log('Not in Teams, running in debug mode');
  const debugEnv = new DebugEnv();
  debugEnv.start();
});
