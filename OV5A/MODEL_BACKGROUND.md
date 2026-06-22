# ATMOS IDEF0 Model Background

## Architecture package

Architecture name: ATMOS

DoDAF model / package: v1

Viewpoint: Systems Architect

Mission / capability:
Air-focused micro-weather context generation, exchange, and maintenance.

Intended audience:
DEVCOM.

Classification / distribution:
Distribution Statement C, controlled distribution.

## Top-level activity

OV-5a top activity:
Conduct Air-Focused Weather Exploitation Operations

Capability name:
ATMOS

Performer:
ATMOS Edge Node

System:
ABLE-LBM

Data product:
Federated Weather Context (Emergent COWP)

## Purpose

The ATMOS operational architecture describes how distributed platforms:

* Observe atmospheric conditions.
* Estimate local micro-weather state.
* Federate and share results.
* Support aviation safety, survivability, and mission effectiveness.

## Viewpoint

Operational architecture described using DoDAF OV-1 through OV-5b.

## Model boundary

Inside the model:

* Observation ingestion.
* Local weather estimation.
* Uncertainty quantification.
* Federation / COWP formation.
* Dissemination of weather context.

Outside the model:

* Route selection.
* Maneuver decisions.
* Command authority.
* Mission execution decisions.

## Critical modeling constraints

COWP is:

* Emergent.
* Not a node.
* Not a decision-maker.

ATMOS:

* Provides descriptive weather context only.
* Does not perform operational decisions.
