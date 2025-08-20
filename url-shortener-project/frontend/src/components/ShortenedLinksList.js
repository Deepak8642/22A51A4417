import React from 'react';
import PropTypes from 'prop-types';
import { List, ListItem, ListItemText, Typography } from '@mui/material';

const ShortenedLinksList = ({ links }) => {
    return (
        <div>
            <Typography variant="h6" gutterBottom>
                Shortened Links
            </Typography>
            <List>
                {links.map((link) => (
                    <ListItem key={link.shortcode}>
                        <ListItemText
                            primary={
                                <a href={link.shortLink} target="_blank" rel="noopener noreferrer">
                                    {link.shortLink}
                                </a>
                            }
                            secondary={`Original URL: ${link.originalUrl} | Expiry: ${new Date(link.expiry).toLocaleString()}`}
                        />
                    </ListItem>
                ))}
            </List>
        </div>
    );
};

ShortenedLinksList.propTypes = {
    links: PropTypes.arrayOf(
        PropTypes.shape({
            shortcode: PropTypes.string.isRequired,
            shortLink: PropTypes.string.isRequired,
            originalUrl: PropTypes.string.isRequired,
            expiry: PropTypes.string.isRequired,
        })
    ).isRequired,
};

export default ShortenedLinksList;