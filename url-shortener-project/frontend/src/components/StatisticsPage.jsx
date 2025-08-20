import React, { useEffect, useState } from 'react';
import { fetchAllStatistics } from '../utils/api';

import { Container, Typography, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Paper } from '@mui/material';

const StatisticsPage = () => {
    const [statistics, setStatistics] = useState([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);

    useEffect(() => {
        const fetchData = async () => {
            try {
                // Use the correct function name
                const data = await fetchAllStatistics();
                setStatistics(data);
            } catch (err) {
                setError(err.message || 'Failed to fetch statistics');
            } finally {
                setLoading(false);
            }
        };

        fetchData();
    }, []);

    if (loading) {
        return <Typography variant="h6">Loading...</Typography>;
    }

    if (error) {
        return <Typography variant="h6" color="error">Error: {error}</Typography>;
    }

    return (
        <Container>
            <Typography variant="h4" gutterBottom>
                Short URL Statistics
            </Typography>
            <TableContainer component={Paper}>
                <Table>
                    <TableHead>
                        <TableRow>
                            <TableCell>Shortened URL</TableCell>
                            <TableCell>Original URL</TableCell>
                            <TableCell>Creation Date</TableCell>
                            <TableCell>Expiry Date</TableCell>
                            <TableCell>Click Count</TableCell>
                        </TableRow>
                    </TableHead>
                    <TableBody>
                        {statistics.map((stat) => (
                            <TableRow key={stat.shortLink}>
                                <TableCell>{stat.shortLink}</TableCell>
                                <TableCell>{stat.originalUrl}</TableCell>
                                <TableCell>{new Date(stat.creationDate).toLocaleString()}</TableCell>
                                <TableCell>{new Date(stat.expiryDate).toLocaleString()}</TableCell>
                                <TableCell>{stat.clickCount}</TableCell>
                            </TableRow>
                        ))}
                    </TableBody>
                </Table>
            </TableContainer>
        </Container>
    );
};

export default StatisticsPage;
