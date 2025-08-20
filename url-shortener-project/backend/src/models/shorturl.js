const mongoose = require('mongoose');

const shortUrlSchema = new mongoose.Schema({
    originalUrl: {
        type: String,
        required: true,
        validate: {
            validator: function(v) {
                const urlRegex = /^(ftp|http|https):\/\/[^ "]+$/;
                return urlRegex.test(v);
            },
            message: props => `${props.value} is not a valid URL!`
        }
    },
    shortCode: {
        type: String,
        required: true,
        unique: true,
        match: /^[a-zA-Z0-9]{1,10}$/ // Alphanumeric and reasonable length
    },
    expiryDate: {
        type: Date,
        required: true
    },
    clickStats: [{
        timestamp: {
            type: Date,
            default: Date.now
        },
        referrer: String,
        geo: String
    }]
}, { timestamps: true });

module.exports = mongoose.model('ShortUrl', shortUrlSchema);