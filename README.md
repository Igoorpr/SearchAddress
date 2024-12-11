## About

SEARCH ADDRESS

Project developed in 2023 with the objective of querying a ZIP Code API and generating an Excel file containing some data returned by the query. The extracted information includes:

- ZIP Code
- Street Address
- Additional Information
- Neighborhood
- State (Federative Unit)
- State
- DDD (Direct Distance Dialing)

The program includes a method that uses regular expressions (regex) to remove accents from the words, ensuring that the data is processed before being exported to Excel.

The API used for ZIP Code queries is [ViaCEP](https://viacep.com.br/).

## Technologies Used

- C#
- .NET Framework 4.8
- Visual Studio
- ViaCep
