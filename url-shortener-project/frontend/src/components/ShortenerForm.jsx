import React, { useState } from 'react';
import { createShortUrl } from '../utils/api';
import { TextField, Button, Typography, Container } from '@mui/material';

const ShortenerForm = () => {
    const [originalUrl, setOriginalUrl] = useState('');
    const [expiryDate, setExpiryDate] = useState('');
    const [shortUrl, setShortUrl] = useState('');
    const [error, setError] = useState('');

    const handleSubmit = async (e) => {
        e.preventDefault();
        try {
            const data = await createShortUrl(originalUrl, expiryDate);
            setShortUrl(data.shortLink);
            setError('');
        } catch (err) {
            setError(err.message);
        }
    };

    return (
        <Container>
            <form onSubmit={handleSubmit}>
                <TextField
                    label="Original URL"
                    value={originalUrl}
                    onChange={(e) => setOriginalUrl(e.target.value)}
                    fullWidth
                    margin="normal"
                />
                <TextField
                    label="Expiry Date"
                    type="date"
                    value={expiryDate}
                    onChange={(e) => setExpiryDate(e.target.value)}
                    fullWidth
                    margin="normal"
                    InputLabelProps={{ shrink: true }}
                />
                <Button type="submit" variant="contained" color="primary">
                    Create Short URL
                </Button>
            </form>

            {shortUrl && (
                <Typography variant="h6" style={{ marginTop: '20px' }}>
                    Short URL: {shortUrl}
                </Typography>
            )}

            {error && (
                <Typography variant="h6" color="error" style={{ marginTop: '20px' }}>
                    Error: {error}
                </Typography>
            )}
        </Container>
    );
};

export default ShortenerForm;
