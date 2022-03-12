# flashtas

![200% unsafe](https://img.shields.io/badge/unsafe-200%25-blue.svg)

An automated input harness for Adobe Flash Player.

This allows you to launch Flash Player with a given SWF file and a prepared set
of inputs which will be forwarded into the program. This requires a working
ActiveX Flash install on the system.

Such software may be useful for tool-assisted speedrunning, or TAS. Note that
rerecording will almost certainly *not* be possible with Flash Player.

## Usage

```cargo run --bin=flashtas -- <path to SWF> <path to input file>```

## Input format

The input format consists of a JSON array-of-objects, where each object is
deserializable to `flashtas_format::AutomatedEvent`.

## Current status

FlashTAS is capable of injecting perfectly-synchronized input to SWFs that are
"movie-like": that is, they have a root timeline with multiple frames that is
never paused. This is because we currently use the root timeline as a sync
mechanism; a paused SWF will thus desynchronize.

Currently, only mouse movement, clicks, and releases can be automated.

## Roadmap

It may be possible in the future to support `fscommand` as a means to
synchronize paused/one-frame movies.

Support for consuming FlashTAS input files is planned to be added to Ruffle.
The input format will adapt and evolve as Ruffle's automated testing needs
change.

## Internals

This software makes heavy use of COM/OLE or "ActiveX technologies", which
allows embedding compatible software in other applications. Flash Player is one
of those applications.

We make use of several Flash Player specific interfaces and quirks, which
prohibits using FlashTAS as a generic ActiveX TASing container.

This workspace also ships with `com-rs` interface definitions for Flash Player,
an OLE Automation type library exporter, and several other core Windows COM
interfaces necessary to set up an OLE container. These wrappers exist primarily
because the COM support in `windows-rs` is not yet as capable as `com-rs`
(though it is getting there). This also means that there is a disgusting amount
of `transmute` calls to bridge between the two COM representations we're stuck
with.