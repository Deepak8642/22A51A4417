const express = require('express');
const router = express.Router();
const shorturlController = require('../controllers/shorturlController');
const loggingMiddleware = require('../middleware/loggingMiddleware');

// Create Short URL
router.post('/shorturls', loggingMiddleware, shorturlController.createShortUrl);

// Retrieve Short URL Statistics
router.get('/shorturls/:shortcode', loggingMiddleware, shorturlController.getShortUrlStatistics);

module.exports = router;