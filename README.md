# OOFAutoUpdater

C# Console application that should be added as an Azure Web Job to update Out of Office for a user based on a custom schedule

Allows a users out of office to automatically switch on when they are not working based on custom working hours each day. No interaction required by the user.

The job should be run once per day and will schedule Exchange OOO to be enabled at a time later that day until the start time the following day.

Each day can have a custom start time and end time and days can be skipped by changing the days parameter for that day
