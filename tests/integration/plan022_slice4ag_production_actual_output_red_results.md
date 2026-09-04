# Plan 022 Slice 4ag reusable Production actual output RED

Date: 2026-08-26

The focused contract was added before implementation and reported `0/6`.
Production Run - List currently labels exact identity with clipped **System
Key**/**Inventory ID** headers. More importantly, the reusable completion path
ignores the operator's **Real Output** entry, creates inventory at the planned
scaled quantity, and always rebuilds the **Last** column blank.

This is meaningful behavioral RED: the public form and packaged harness are
available, and the failure matches the visible completed-run checkpoint.
