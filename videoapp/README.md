# Introduction

This project demonstracts how to use `microsoft/teams-js` to develop a Teams video app.

# To develop a Teams video app

## Use microsoft/teams-js SDK

Please see "videoapp\src\main.ts" as an example.
Every SDK API call should happen after initialization:
```javascript
import { app, videoEffects } from '@microsoft/teams-js';

app
  .initialize([`${your app domain here}`])
  .then(() => {
    initializeVideoApp();
  })
  .catch(() => {
    // error handling
  });
```
After initialization, register callbacks for receiving video effect changed event and video frames sent from Teams client.
### To receive video effect changed event

Use `registerForVideoEffect` to register callback to get the current selected effect id.

```javascript
type VideoEffectCallback = (effectId: string | undefined) => Promise<void>;

videoEffects.registerForVideoEffect(callback: VideoEffectCallback);

```
On `VideoEffectCallback` receiving an efffect id, the video app should do the preparation, if any, on receiving the effect id. When the app finishes the preparation and is ready to process the new effect, resolve the promise.

There are 2 ways that the effect id gets updated: 
1. The user selects an effect from Teams client。
2. The user selects an effect from the video app page: in this case, the app should call `videoEffects.notifySelectedVideoEffectChanged` to notify the new effect id - see `initializeVideoApp` in main.ts to know how use this API.

### To receive and process video frames

Use `registerForVideoFrame` to register callback to receive and process video frames, it takes 2 callbacks to receive video frames sent by either classic Teams or the new Teams, please see the [doc](https://github.com/OfficeDev/microsoft-teams-library-js/blob/4976ceee315840260a687c235d69076b5c1d135a/packages/teams-js/src/public/videoEffects.ts#L186) here to understand how to use this API.

# Prepare a manifect package

Please see the `./menifest` folder.
You will need to prepare thumbnails for each video effect. Name the thumbnail as `${effectId}_thumbnail.png`. 2 images for the app icon are required too, name them `${appId}_{large|small}Image.png`.
Prepare the `manifest.json`, you can copy the one in the manifest folder and do the corresponding changes:
* Use the domain of your app in `validDomains` section.
* List all the video effects provided by your app in `meetingExtensionDefinition.videoFilters` section. Please note the naming convention, name each video effect as `${category}_effect name`, category can be `Other|Frame|other video app`.
* Update `videoFiltersConfigurationUrl` to your app's full address.

Once the files are ready, pack them together into a zip file, all the files should be under the root path, then sideload the app following this [artical](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload). You should be able to run your app on Teams now.
Please start a meeting, open the "Effects and avatars" pane, click "more effects" and find your video effects.
