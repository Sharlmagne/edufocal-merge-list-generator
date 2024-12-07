# MyWPFApp

![License](https://img.shields.io/badge/license-MIT-green)
![.NET Version](https://img.shields.io/badge/.NET-8.0-blue)

## Overview

**Edufocal Merge List Generator** is a simple WPF application that solves a privacy issue where employee information needs to be secured. The application generates a list of employees with their respective information by combining two separate lists. The first list contains the employees' sensitive information, while the second list contains the employees' aliases along with all the information required to mail merge a certificate. The application combines the two lists to generate a single list with the employees' sensitive information and the information required to mail merge a certificate, while keeping the sensitive information secure on the user's machine.

## Features

- **MVVM Architecture**: Implements the Model-View-ViewModel design pattern for clean separation of concerns.
- **Dependency Injection**: Uses .NET's built-in DI container for managing dependencies.
- **Settings Management**: Persistent settings using `ConfigurationManager`.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
- Visual Studio 2022 (or newer) with the **.NET Desktop Development** workload
- Windows 10 or later

## Getting Started

1. Clone the repository:
   ```bash
   git clone https://github.com/Sharlmagne/edufocal-merge-list-generator.git
   cd edufocal-merge-list-generator
   ```
2. Open the solution and restore the dependencies.
    ```bash 
   dotnet restore
    ```
3. Build the solution.
4. Run the application.

## Download the application
https://sharlmagne.github.io/edufocal-merge-list-generator/EdufocalMergeListGenerator.application