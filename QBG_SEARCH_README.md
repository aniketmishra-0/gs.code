# QBG Search Tool - User Guide

## Overview

The QBG Search Tool is a modern, feature-rich search interface designed for educational institutions. It provides powerful search capabilities with an intuitive user interface, dark mode support, favorites management, and advanced filtering options.

## Features

### 🔍 Advanced Search
- Real-time search with instant results
- Mock data included for demonstration
- Easy integration with Google Apps Script backend

### ⭐ Search History & Favorites
- **Save Favorites**: Click the star icon to save any search with a custom name
- **Quick Access**: Instantly reload saved searches from the Favorites panel
- **Recent History**: Automatically tracks your last 20 searches with timestamps
- **Manage Items**: Delete individual favorites or history items with one click
- **Persistent Storage**: All favorites and history are saved in browser LocalStorage

### 🌙 Dark Mode
- **Toggle Button**: Switch between light and dark themes with one click
- **System Detection**: Automatically detects your system's dark mode preference
- **Smooth Transitions**: Beautiful animations when switching themes
- **Complete Coverage**: Every UI element is styled for both themes
- **Persistent**: Your theme choice is saved and remembered

### 🔍 Advanced Filters
- **Exclude Terms**: Filter out unwanted results (use `-` prefix)
- **Whole Words**: Match only complete words
- **Case Sensitive**: Enable case-sensitive searching
- **Visual Indicators**: Active filters are displayed as tags
- **Saved with Favorites**: Filter settings are saved when you favorite a search

### 💡 Autocomplete & Suggestions
- **Real-time Suggestions**: See suggestions as you type
- **Popular Terms**: Predefined popular search terms
- **Recent Searches**: Your recent searches appear in suggestions
- **Keyboard Navigation**: Use arrow keys to navigate, Enter to select, Esc to close
- **Highlighted Matches**: See which part of the suggestion matches your query

## How to Use

### Basic Search
1. Open `QBGSearchTool.html` in any modern web browser
2. Type your search query in the search box
3. Click "Search" or press Enter
4. View results below

### Using Autocomplete
1. Start typing in the search box
2. Suggestions will appear automatically
3. Use ↑/↓ arrow keys to navigate suggestions
4. Press Enter or click to select a suggestion
5. Press Esc to close suggestions

### Saving Favorites
1. Enter a search query
2. Click the ⭐ star icon next to the search button
3. Enter a name for your favorite
4. Click "Save"
5. Access your favorites from the "Favorites & History" panel

### Viewing History
1. Click "⭐ Favorites & History" button
2. Switch to "🕒 Recent Searches" tab
3. Click any search to reload it
4. Click 🗑️ to delete an item

### Using Advanced Filters
1. Click "🔍 Advanced Filters" to expand the panel
2. Check the filters you want to apply:
   - **Exclude terms**: Removes results containing specified terms
   - **Match whole words**: Only matches complete words
   - **Case sensitive**: Respects letter case in searches
3. Active filters appear as tags below
4. Filters are automatically applied to searches

### Switching Themes
1. Click the "🌙 Dark Mode" button in the header
2. The theme switches immediately with smooth animations
3. Your preference is saved automatically

## Technical Details

### Browser Compatibility
- Chrome/Edge 88+
- Firefox 78+
- Safari 14+
- Opera 74+

### Storage
- Uses browser LocalStorage for persistence
- Favorites: Unlimited (browser-dependent, typically 5-10MB)
- History: Last 20 searches

### Mobile Support
- Fully responsive design
- Touch-friendly interface
- Optimized layouts for small screens

### Accessibility
- ARIA labels on all interactive elements
- Keyboard navigation support
- Semantic HTML structure
- Screen reader compatible

## Integration with Google Apps Script

The search tool is designed to work with a Google Apps Script backend. To integrate:

1. Replace the `mockSearch()` function in the JavaScript code with an actual backend call:

```javascript
function performSearch() {
    const query = elements.searchInput.value.trim();
    if (!query) return;
    
    // Add to history
    addToHistory(query);
    
    // Show loading state
    elements.resultsCount.innerHTML = '<span class="loading"></span> Searching...';
    elements.resultsContainer.innerHTML = '';
    
    // Call Google Apps Script backend
    google.script.run
        .withSuccessHandler(displayResults)
        .withFailureHandler(handleError)
        .searchData(query, state.filters);
}
```

2. Update the autocomplete data with real terms from your system:

```javascript
state.autocompleteData = [
    // Your actual search terms here
];
```

3. Create corresponding server-side function in Google Apps Script:

```javascript
function searchData(query, filters) {
    // Your search logic here
    return results;
}
```

## Customization

### Colors
Edit the CSS variables in the `:root` and `[data-theme="dark"]` sections to customize colors:

```css
:root {
    --primary-color: #4f46e5;  /* Change primary color */
    --background: #ffffff;      /* Light mode background */
    /* ... more variables */
}
```

### Autocomplete Terms
Edit the `state.autocompleteData` array to add your own suggestions:

```javascript
autocompleteData: [
    'your search term 1',
    'your search term 2',
    // ... more terms
]
```

### Mock Search Results
Edit the `mockData` array in `mockSearch()` function to customize demo results.

## Troubleshooting

### Favorites/History not persisting
- Ensure cookies and LocalStorage are enabled in your browser
- Check browser privacy settings
- Try clearing LocalStorage: `localStorage.clear()`

### Autocomplete not showing
- Ensure you're typing at least one character
- Check browser console for JavaScript errors
- Verify the `autocompleteData` array has items

### Dark mode not working
- Check browser compatibility
- Ensure JavaScript is enabled
- Clear browser cache

## Support

For issues or questions:
1. Check browser console for errors
2. Verify browser compatibility
3. Ensure JavaScript is enabled
4. Test in incognito/private mode to rule out extensions

## License

This tool is designed for educational use in school management systems.

## Version

Version 1.0.0 - November 2025
