# Introduction

This document demonstrates how to use the `microsoft/teams-js` SDK to develop a Teams video app.

# Developing a Teams Video App

## Using the `microsoft/teams-js` SDK

Refer to `videoapp/src/main.ts` for an example. Ensure that every SDK API call occurs after initialization:

```javascript
import { app, videoEffects } from '@microsoft/teams-js';

app
  .initialize([`${your app domain here}`])
  .then(() => {
    initializeVideoApp();
  })
  .catch(() => {
    // Handle errors here
  });
```

After initialization, register callbacks to handle video effect changes and video frames sent from the Teams client.

### Receiving Video Effect Changes

To receive video effect change events, use `registerForVideoEffect` to register a callback that receives the current selected effect ID:

```javascript
type VideoEffectCallback = (effectId: string | undefined) => Promise<void>;

videoEffects.registerForVideoEffect(callback: VideoEffectCallback);
```

When the `VideoEffectCallback` receives an effect ID, the video app should prepare to handle the new effect. Once the app is ready to process the effect, resolve the promise.

Effect IDs can be updated in two ways:
1. The user selects an effect from the Teams client.
2. The user selects an effect from the video app page. In this case, the app should call `videoEffects.notifySelectedVideoEffectChanged` to notify the new effect ID. Refer to `initializeVideoApp` in `main.ts` for details on using this API.

### Receiving and Processing Video Frames

To receive and process video frames, use `registerForVideoFrame` and provide two callbacks to handle video frames from either the classic Teams client or the new Teams client. For more information, refer to the [documentation](https://github.com/OfficeDev/microsoft-teams-library-js/blob/4976ceee315840260a687c235d69076b5c1d135a/packages/teams-js/src/public/videoEffects.ts#L186).

# Preparing the Manifest Package

Refer to the `./manifest` folder. 

You need to prepare thumbnails for each video effect, naming them as `${effectId}_thumbnail.png`. Additionally, provide two images for the app icon, named `${appId}_{large|small}Image.png`. 

Prepare the `manifest.json` file by copying the one in the manifest folder and making the following changes:
- Update the `validDomains` section with your app's domain.
- List all video effects provided by your app in the `meetingExtensionDefinition.videoFilters` section. Follow the naming convention `${category}_effect name`, where `category` can be `Other`, `Frame`, or `${video app name}`.
- Set `videoFiltersConfigurationUrl` to your app's full address.

Once you have prepared these files, zip them together with all files at the root level. Follow the instructions in this [article](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload) to sideload the app. You should then be able to run your app in Teams. Start a meeting, open the "Effects and avatars" pane, click "More effects," and find your video effects.