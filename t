<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Voice Controlled Object and Text Recognition</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            text-align: center;
            margin-top: 20px;
        }
        #video {
            border: 1px solid black;
        }
        #controls {
            margin: 10px;
        }
        #controls button {
            padding: 10px 20px;
            font-size: 16px;
            margin: 5px;
            cursor: pointer;
        }
    </style>
</head>
<body>
    <h1>Voice Controlled Object and Text Recognition</h1>
    <p>Please grant camera and microphone permissions when prompted.</p>
    <video id="video" width="600" height="400" autoplay muted></video>
    <canvas id="canvas" style="display: none;"></canvas>
    <div id="controls">
        <button id="toggleButton">Start Camera</button>
    </div>

    <!-- Load TensorFlow.js, COCO-SSD model, and Tesseract.js -->
    <script src="https://cdn.jsdelivr.net/npm/@tensorflow/tfjs"></script>
    <script src="https://cdn.jsdelivr.net/npm/@tensorflow-models/coco-ssd"></script>
    <script src="https://cdn.jsdelivr.net/npm/tesseract.js@2.1.1/dist/tesseract.min.js"></script>

    <script>
        const video = document.getElementById('video');
        const canvas = document.getElementById('canvas');
        const context = canvas.getContext('2d');
        const toggleButton = document.getElementById('toggleButton');
        let model, detectionInterval;
        let mode = null;
        let isDetecting = false;
        let lastDetectedObjects = [];
        let isSpeaking = false;
        let lastDetectionTime = 0;
        const detectionIntervalMs = 2000; // Shorter interval for better accuracy

        // Initialize speech recognition
        const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
        recognition.continuous = true;
        recognition.lang = 'en-US';

        // Handle speech recognition results
        recognition.onresult = (event) => {
            const command = event.results[event.results.length - 1][0].transcript.trim().toLowerCase();
            handleVoiceCommand(command);
        };

        recognition.onerror = (event) => {
            console.error('Speech recognition error:', event.error);
        };

        recognition.onend = () => {
            recognition.start(); // Restart recognition
        };

        recognition.start();

        // Handle voice commands
        function handleVoiceCommand(command) {
            console.log('Command received:', command);

            if (command.includes('start camera')) {
                startVideo();
                toggleButton.textContent = 'Stop Camera';
                stopDetection();
                isDetecting = false; // Ensure detecting is stopped
            } else if (command.includes('stop camera')) {
                stopVideo();
                toggleButton.textContent = 'Start Camera';
                stopDetection();
                isDetecting = false; // Ensure detecting is stopped
            } else if (command.includes('what is it')) {
                if (!isDetecting) {
                    mode = 'objects';
                    isDetecting = true;
                    startDetection();
                }
            } else if (command.includes('read')) {
                if (!isDetecting) {
                    mode = 'text';
                    isDetecting = true;
                    startDetection();
                }
            }
        }

        // Stop detection
        function stopDetection() {
            if (detectionInterval) {
                clearInterval(detectionInterval);
                detectionInterval = null;
            }
        }

        // Start detection loop
        function startDetection() {
            if (detectionInterval) {
                stopDetection(); // Ensure no multiple intervals
            }
            detectionInterval = setInterval(() => {
                if (isDetecting) {
                    detectFrame();
                }
            }, detectionIntervalMs);
        }

        // Load the COCO-SSD model
        async function loadModel() {
            model = await cocoSsd.load();
            console.log('Model loaded');
        }

        // Start the video feed
        function startVideo() {
            navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" }, audio: true })
                .then((stream) => {
                    video.srcObject = stream;
                    video.addEventListener('loadeddata', () => {
                        canvas.width = video.videoWidth;
                        canvas.height = video.videoHeight;
                    });
                })
                .catch((err) => {
                    console.error('Error accessing webcam:', err);
                    alert('Please grant camera and microphone permissions for this feature to work.');
                });
        }

        // Stop the video feed
        function stopVideo() {
            const stream = video.srcObject;
            if (stream) {
                stream.getTracks().forEach((track) => track.stop());
                video.srcObject = null;
            }
        }

        // Detect objects or text
        async function detectFrame() {
            const currentTime = Date.now();
            if (isDetecting && (currentTime - lastDetectionTime) > detectionIntervalMs) {
                lastDetectionTime = currentTime;
                context.drawImage(video, 0, 0, canvas.width, canvas.height);

                if (mode === 'objects') {
                    await detectObjects();
                } else if (mode === 'text') {
                    await detectText();
                }
                isDetecting = false; // Stop detection after processing
            }
        }

        // Detect objects using COCO-SSD
        async function detectObjects() {
            const predictions = await model.detect(video);
            drawPredictions(predictions);

            const detectedObjects = [...new Set(predictions.map(pred => pred.class))];

            if (detectedObjects.length > 0 && !arraysEqual(detectedObjects, lastDetectedObjects)) {
                lastDetectedObjects = detectedObjects;
                speak(detectedObjects.join(', '));
            }
        }

        // Draw bounding boxes and labels on the canvas
        function drawPredictions(predictions) {
            context.clearRect(0, 0, canvas.width, canvas.height);
            predictions.forEach(pred => {
                context.beginPath();
                context.rect(...pred.bbox);
                context.lineWidth = 2;
                context.strokeStyle = 'red';
                context.fillStyle = 'red';
                context.stroke();
                context.fillText(`${pred.class} (${Math.round(pred.score * 100)}%)`, pred.bbox[0], pred.bbox[1] > 10 ? pred.bbox[1] - 5 : 10);
            });
        }

        // Detect text using Tesseract.js
        async function detectText() {
            const { data: { text } } = await Tesseract.recognize(canvas, 'eng', {
                logger: info => console.log(info),
                corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js@2.1.1/dist/tesseract-core.wasm.js'
            });
            if (text.trim()) {
                speak(text.trim());
            } else {
                speak('No text detected');
            }
        }

        // Speak the detected objects or text with increased volume and clarity
        function speak(text) {
            if (isSpeaking) {
                speechSynthesis.cancel();
            }
            const utterance = new SpeechSynthesisUtterance(text);
            isSpeaking = true;
            utterance.onend = () => {
                isSpeaking = false;
            };
            utterance.onerror = (event) => {
                isSpeaking = false;
                console.error('Speech synthesis error:', event.error);
            };
            utterance.pitch = 1.5;  // Increase pitch for clarity
            utterance.rate = 1.2;   // Increase rate for faster speech
            utterance.volume = 1;   // Max volume
            speechSynthesis.speak(utterance);
        }

        // Check if two arrays are equal
        function arraysEqual(a, b) {
            if (a.length !== b.length) return false;
            for (let i = 0; i < a.length; i++) {
                if (a[i] !== b[i]) return false;
            }
            return true;
        }

        // Initialize model and start video when page is ready
        document.addEventListener('DOMContentLoaded', async () => {
            await loadModel();
            toggleButton.addEventListener('click', () => {
                if (video.srcObject) {
                    stopVideo();
                    toggleButton.textContent = 'Start Camera';
                    stopDetection();
                    isDetecting = false;
                } else {
                    startVideo();
                    toggleButton.textContent = 'Stop Camera';
                    stopDetection();
                    isDetecting = false;
                }
            });
        });
    </script>
</body>
</html>
