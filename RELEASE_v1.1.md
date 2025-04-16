# Quizizz AI Import v1.1

We're excited to announce version 1.1 of Quizizz AI Import with significant improvements to handling large documents and Vietnamese content!

## What's New

- **Large Document Support**: Process documents of any size with new chunking functionality
  - Automatically splits large files into manageable chunks
  - Progress indicator shows processing status
  - Configurable chunk size for balancing speed and reliability

- **Enhanced Vietnamese Support**: Better detection and formatting of Vietnamese questions
  - Improved parsing of "Câu" numbered questions
  - Maintains proper Vietnamese text encoding
  - Better handling of question/answer structures

- **UI Improvements**:
  - Notification system for important updates and announcements
  - More responsive design with animated transitions
  - Better dark mode support

- **Export Enhancements**:
  - Native file saving through Tauri APIs
  - Better Excel formatting for Quizizz import
  - Support for adding images, time limits, and explanations

## Bug Fixes

- Fixed issues with skipping questions during parsing
- Improved error handling for API failures
- Better handling of special characters in Vietnamese text
- More robust question detection patterns

## Technical Changes

- Updated to latest Gemini AI API
- Optimized for better performance with large documents
- Improved codebase organization and error handling

Thank you for using Quizizz AI Import! Please report any issues or feature requests on our GitHub repository.