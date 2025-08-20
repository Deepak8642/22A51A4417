const express = require('express');
const mongoose = require('mongoose');
const shortUrlRoutes = require('./routes/shorturls');
const loggingMiddleware = require('./middleware/loggingMiddleware');

const app = express();
app.use(express.json());
app.use(loggingMiddleware);

app.get('/', (req, res) => {
    res.send('URL Shortener Service is running');
});

app.use('/', shortUrlRoutes);

const mongoUri = 'mongodb://localhost:27017/urlshortener';
mongoose.connect(mongoUri, { useNewUrlParser: true, useUnifiedTopology: true })
    .then(() => {
        app.listen(5000, () => {});
    })
    .catch((err) => {
        loggingMiddleware.logError(err);
        process.exit(1);
    });