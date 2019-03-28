# VBScript based VSTS result updater

This is a simple implementation of test execution result updation activity by extending the REST API's provided and supported by Microsoft VSTS. 

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### WHAT DOES THIS CODE DO? 
```
1. This will get the Project information based on test case ID, suite ID, Plan ID. 
2. Create a run in that particular suite ID provided
3. Update the result for that particular run. 
4. Add attachments to the particular run.
5. If the test case status is marked as FAILED, it creates a bug and links the test case run to the Bug.
```

### Prerequisites

What things you need to install the software and how to install them

1. VBScript - Available as part of Windows OS
2. [Microsoft VSTS API reference](https://docs.microsoft.com/en-us/rest/api/azure/devops/)


## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.


## Authors

* **Aravinth Rajan** - *Initial work* - [AravinthR](https://github.com/AravinthR)
