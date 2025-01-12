# VBScript IsEmpty Function Inconsistency

This example demonstrates a potential issue with VBScript's `IsEmpty` function when dealing with Variants that may hold empty strings ("") or Null values.  The `IsEmpty` function doesn't consistently distinguish between these two states, which can lead to unexpected behavior in your scripts.

The `bug.vbs` file showcases the problematic behavior. The solution (`bugSolution.vbs`) provides a more robust approach using explicit checks for both empty strings and Null values.