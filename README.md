# VSTSresultUpdater

This is a simple implementation of REST based test result updation activity by extending the REST API's provided and supported by Microsoft VSTS. 
The reason behind this implementation is to reduce the human internenvtion during the test phase and for teams who are completely migrating to DevOps based approach where the testing activities are pipelined. 

=-=-=-=-=-=-=-=-=-=- WHAT DOES THIS CODE DO =-=-=-=-=-=-=-=-=-=-=-=-=-
1. This will get the Project information based on test case ID, suite ID, Plan ID. 
2. Create a run in that particular suite ID provided
3. Update the result for that particular run. 
4. Add attachments to the particular run.
5. If the test case status is marked as FAILED, it creates a bug and links the test case run to the Bug.
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=
