import React, { useState } from 'react'
import { Container, Typography, Box, Tabs, Tab } from '@mui/material'
import ShortenerForm from './components/ShortenerForm'
import ShortenedLinksList from './components/ShortenedLinksList'
import StatisticsPage from './components/StatisticsPage'

function App() {
    const [tab, setTab] = useState(0)
    const [shortenedLinks, setShortenedLinks] = useState([])

    return (
        <Container maxWidth="md">
            <Box sx={{ my: 4 }}>
                <Typography variant="h4" align="center" gutterBottom>
                    URL Shortener
                </Typography>
                <Tabs value={tab} onChange={(_, v) => setTab(v)} centered>
                    <Tab label="Shorten URLs" />
                    <Tab label="Statistics" />
                </Tabs>
                {tab === 0 && (
                    <>
                        <ShortenerForm setShortenedLinks={setShortenedLinks} />
                        <ShortenedLinksList links={shortenedLinks} />
                    </>
                )}
                {tab === 1 && <StatisticsPage links={shortenedLinks} />}
            </Box>
        </Container>
    )
}

export default App