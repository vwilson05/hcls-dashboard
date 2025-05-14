# Healthcare Delivery Dashboard

A Streamlit-based dashboard for healthcare delivery leadership to monitor projects, pipeline, risks, staffing, and strategic opportunities.

## Features

- 📊 Real-time data visualization from Google Sheets
- 🔍 Natural language querying using OpenAI's GPT-4
- 📈 Key metrics tracking
- 📱 Responsive and intuitive interface
- 🔐 Secure credential management

## Setup

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up environment variables:
   - Copy `.env.template` to `.env`
   - Fill in your credentials:
     - Google Sheets service account credentials
     - OpenAI API key
     - Google Sheet name

4. Place your Google Sheets service account credentials file (`credentials.json`) in the project root

## Running the App

```bash
streamlit run app.py
```

## Project Structure

- `app.py`: Main Streamlit application
- `requirements.txt`: Python dependencies
- `.env`: Environment variables (not tracked in git)
- `credentials.json`: Google Sheets service account credentials (not tracked in git)

## Data Structure

The dashboard expects the following worksheets in your Google Sheet:

1. Project Inventory
2. Project Risks
3. Pipeline
4. Team Utilization
5. Talent Gaps
6. Operational Gaps
7. Executive Activity
8. Scenario Model Inputs
9. Do Nothing Scenario
10. Proposed Scenario
11. Scenario Comparison

## Security Notes

- Never commit `.env` or `credentials.json` to version control
- Use environment variables for sensitive data
- Consider implementing role-based access control for production

## Future Enhancements (Phase 2)

- Interactive charts and visualizations
- Advanced filtering capabilities
- Sheet editing functionality
- Role-based permissions
- Automated data refresh
- Export capabilities

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details. 