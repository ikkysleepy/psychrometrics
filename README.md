# Psychrometric calculations tool


## Summary

Our Psychrometrics software will allow you to do psychrometric calculations in MS Excel. The software is based on an Open sourced project,PsychroLib, which uses ASHRAE correlations and is more than accurate enough for HVAC applications. It works with both English (IP) and Metric (SI) units.

## Features

The Psychrometrics Excel add-in, allows you to calculate the following properties in any Excel spreadsheet for psychrometric calculations and is compatible with MS Excel 2013 or newer.
 
- Enthalpy (h)
- Dewpoint Temperature (dp)
- Relative Humidity (%)
- Humidity Ratio (W)
- Specific Volume (v)
- Wet Bulb Temperature (wb)


## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 3-9-2020 | Initial release
1.1 | 8-5-2021 | Update to GitHub page hosting

## Scenario: A contextual add-in



## Run the sample from Localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:
    
    ```console
    npm install --global http-server
    ```
    
2. Use a tool such as openssl to generate a self-signed certificate for the web server. Move the cert.pem and key.pem files to the webworker-customfunction folder for this sample.
3. From a command prompt, go to the web-worker folder and run the following command:
    
    ```console
    http-server -S --cors . -p 3000
    ```
    
4. To reroute to localhost run office-addin-https-reverse-proxy. If you haven't installed this you can do this with the following command:
    
    ```console
    npm install --global office-addin-https-reverse-proxy
    ```
    
    To reroute run the following in another command prompt:
    
    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000
    ```
    
5. Follow the steps in Run the sample, but upload the `manifest-localhost.xml` file for step 6.

## Security notes

None

## Copyright

Copyright (c) 2022 kW Engineering. All rights reserved.

