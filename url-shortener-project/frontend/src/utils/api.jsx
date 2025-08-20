import axios from 'axios';

const API_BASE_URL = 'http://localhost:5000/api'; // replace with your backend URL

export const createShortUrl = async (originalUrl, expiryDate) => {
    try {
        const response = await axios.post(`${API_BASE_URL}/shorten`, { originalUrl, expiryDate });
        return response.data;
    } catch (error) {
        throw new Error(error.response?.data?.message || 'Failed to create short URL');
    }
};

export const fetchAllStatistics = async () => {
    try {
        const response = await axios.get(`${API_BASE_URL}/statistics`);
        return response.data;
    } catch (error) {
        throw new Error(error.response?.data?.message || 'Failed to fetch statistics');
    }
};

export const getShortUrlStatistics = async (shortLink) => {
    try {
        const response = await axios.get(`${API_BASE_URL}/statistics/${shortLink}`);
        return response.data;
    } catch (error) {
        throw new Error(error.response?.data?.message || 'Failed to fetch short URL statistics');
    }
};
  