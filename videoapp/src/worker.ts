onmessage = function (e) {
    const videoFrame = e.data as VideoFrame;
    postMessage(videoFrame, [videoFrame]);
}