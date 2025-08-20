const crypto = require('crypto');

function generateShortCode(length = 6) {
    const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    let code = ''
    for (let i = 0; i < length; i++) {
        code += chars.charAt(Math.floor(Math.random() * chars.length))
    }
    return code
}

const isShortcodeUnique = async(shortcode, ShortUrlModel) => {
    const existingUrl = await ShortUrlModel.findOne({ shortcode });
    return !existingUrl;
};

const createUniqueShortcode = async(ShortUrlModel, customShortcode) => {
    if (customShortcode) {
        const isUnique = await isShortcodeUnique(customShortcode, ShortUrlModel);
        if (isUnique) {
            return customShortcode;
        }
        throw new Error('Custom shortcode is already in use.');
    }

    let shortcode;
    do {
        shortcode = generateShortCode();
    } while (!(await isShortcodeUnique(shortcode, ShortUrlModel)));

    return shortcode;
};

module.exports = {
    generateShortCode,
    createUniqueShortcode,
};