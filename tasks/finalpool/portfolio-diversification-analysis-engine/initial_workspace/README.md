# Portfolio Analysis Engine — Working Folder

This workspace contains the current investment portfolio and reference configuration.

Files
- current_portfolio.csv — current holdings with ticker, asset class, sector, weight, and cost basis
- config.json — analysis configuration (risk-free rate, time horizon, constraints)
- data.csv — placeholder reference data (not used directly)

Notes
- The portfolio holdings are equities and equity-proxy ETFs (broad market and commodity ETFs) tradable on standard equity markets, so all instruments can be analyzed using standard equity market data services.
- Use the equity-tradable proxies for the broader asset-class exposure analysis (e.g., GLD for commodity exposure, SPY for broad market beta).
