# Create Exchange Reservable Resource

If you've ever attempted to create a reservable workspace in Exchange, you might've realized that there is not an option to create a Resource of type 'Workspace' in the Exchange Admin Center. Currently, the admin center only supports creating resources of type 'Room' and 'Equipment'. 

## ✨ Features

-   Creates reservable Workspace in Exchange Online

## 📝 Output

None

## 🚀 Getting Started

### Prerequisites

-   Exchange Administrator Role

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
```

## Usage

The intention is for this script to be called by a parent script that will pass in the required parameters. This allows you to run the script against multiple users and potentially multiple tenants.
Below is an example of how you might call the script.

```powershell

```

## 🤝 Contributing

Contributions, issues and feature requests are welcome!
