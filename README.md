# OSADPConnector
Open-source Datacap-ADP Connector

This is a set of custom Datacap actions that provide a bidirectional connector from Datacap to IBM's Automation Document Processing (ADP), allowing Datacap customers to send documents to ADP for classification and extraction, then get the results back in Datacap for further processing. This connector allows customers for the first time to leverage both the power and flexibility of Datacap as well as the simple configuration and AI power of ADP.

See "The ADPDemo Sample Application.pdf" for more information about the sample Datacap application, which includes the connector DLL. If you want to build the code, you can build it in Visual Studio, but you'll need to install the Datacap SDK, available at https://github.com/ibm-ecm/datacap-developer-kit.

# Overview
Datacap is IBM's traditional document capture solution. It has been on the market for decades, has a large install base, and provides a great deal of functionality, though it can be difficult to configure and to learn how to build applications with it.

ADP is IBM's new AI-based document capture solution. It's only been on the market for a few years, provides a no-code set-up experience, automatically classifies and categorizes documents, extracts data, and uses AI and deep learning throughout the process.

With the normal out-of-the-box integration between Datacap and ADP, documents are ingested into Datacap, then sent to ADP for all subsequent processing. This leverages the strengths of ADP, but users aren't able to take advantage of the rich functionality of Datacap.
![image](https://github.com/IBM/Datacap-ADP-Connector/assets/40502969/d57a0b7c-7d4d-461d-890a-d8d3c35c0410)

With this connector, you can get the best of both worlds -- the easy configuration and AI-based classification extraction that ADP provides, plus the powerful validation, lookups and delivery in Datacap. You can also optionally leverage the benefits provided by the Datacap Accelerator (for more information on that, see your IBM Expert Labs team).
![image](https://github.com/IBM/Datacap-ADP-Connector/assets/40502969/4ce1748e-7e5e-4c32-8f07-7c4d0ca86bb7)
