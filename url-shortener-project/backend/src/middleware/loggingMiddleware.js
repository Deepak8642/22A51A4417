const loggingMiddleware = (req, res, next) => {
    const start = Date.now();
    
    // Log the incoming request
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`);

    // Capture the response on finish
    res.on('finish', () => {
        const duration = Date.now() - start;
        console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl} ${res.statusCode} - ${duration}ms`);
    });

    next();
};

module.exports = loggingMiddleware;