const ShortUrl = require('../models/shorturl')
const { generateShortCode } = require('../utils/shortcodeGenerator')
const loggingMiddleware = require('../middleware/loggingMiddleware')

exports.createShortUrl = async(req, res) => {
    const { url, validity, shortcode } = req.body
    const urlPattern = new RegExp('^(https?:\\/\\/)?' +
        '((([a-z\\d]([a-z\\d-]*[a-z\\d])?)\\.)+[a-z]{2,}|' +
        'localhost|' +
        '\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}|' +
        '\\[?[a-fA-F0-9]*:[a-fA-F0-9:]+\\]?)' +
        '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' +
        '(\\?[;&a-z\\d%_.~+=-]*)?' +
        '(\\#[-a-z\\d_]*)?$', 'i')

    if (!urlPattern.test(url)) {
        return res.status(400).json({ error: 'Invalid URL format' })
    }

    const expiryDuration = validity ? validity : 30
    const expiryDate = new Date(Date.now() + expiryDuration * 60000)
    let shortCodeToUse = shortcode
    if (!shortCodeToUse) {
        shortCodeToUse = generateShortCode()
    } else {
        const existingShortUrl = await ShortUrl.findOne({ shortcode: shortCodeToUse })
        if (existingShortUrl) {
            return res.status(400).json({ error: 'Shortcode already in use' })
        }
    }

    const newShortUrl = new ShortUrl({
        url,
        shortcode: shortCodeToUse,
        expiry: expiryDate,
        clicks: 0,
        clickData: []
    })

    try {
        await newShortUrl.save()
        res.status(201).json({
            shortLink: `http://hostname:port/${shortCodeToUse}`,
            expiry: expiryDate.toISOString()
        })
    } catch (error) {
        loggingMiddleware.logError(error)
        res.status(500).json({ error: 'Internal server error' })
    }
}

exports.getShortUrlStatistics = async(req, res) => {
    const { shortcode } = req.params
    try {
        const shortUrl = await ShortUrl.findOne({ shortcode })
        if (!shortUrl) {
            return res.status(404).json({ error: 'Shortcode not found' })
        }
        res.status(200).json({
            originalUrl: shortUrl.url,
            creationDate: shortUrl.createdAt,
            expiryDate: shortUrl.expiry,
            totalClicks: shortUrl.clicks,
            clickData: shortUrl.clickData
        })
    } catch (error) {
        loggingMiddleware.logError(error)
        res.status(500).json({ error: 'Internal server error' })
    }
}