# ALMBulkExecutor
Project to automate different redundant manual tasks in ALM

An automation tool developed to minimize time taken to address administrative activities in ALM.

It uses Java Swings API for the front-end and the OTA COM API of HP for back-end.

The entire coding of the tool has been done using Java and JACOB (Java Com Bridge) to make JNI calls to the COM API of HP-ALM. 

At present, this tool can be used for the following 3 processes:

1. Step by Step test execution(Passed/Failed) of a test case in Test Lab and attaching screenshots.

2. Fast Run execution(Passed/Failed/Blocked) of the test case in Test Lab and attaching screenshot.

3. Mass blocking/unblocking (Block/No run) of test scripts based on test sets in Test Lab and linking defects.


