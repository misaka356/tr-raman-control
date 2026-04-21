# Integrated Controller Usage

## SDK Root

The controller now uses the Andor official SDK directly instead of driving the SOLIS UI.

By default it looks for the vendored runtime here:

`C:\Users\adimn\Desktop\code\andor\vendor\andor_sdk`

You can also pass another SDK directory with `--andor-sdk-root`.

## Example

```powershell
python .\tr_raman_integrated_controller.py `
  --andor-sdk-root "C:\Users\adimn\Desktop\code\andor\vendor\andor_sdk" `
  --connection-type usb `
  --visa-resource "USB0::0x1AB1::0x0643::DGXXXXXXXXX::INSTR" `
  --output-dir C:\AndorOutput `
  --sample-name test `
  --phase-start 0 `
  --phase-stop 360 `
  --phase-step 10 `
  --repeats 3 `
  --ch1-delay 0.002 `
  --ch1-freq 1000 `
  --ch1-amp 1.0 `
  --ch2-freq 1000 `
  --ch2-amp 5.0 `
  --center-wavelength 500 `
  --grating 1 `
  --exposure 0.2 `
  --trigger-mode 1
```

## What It Does

1. load the vendored Andor SDK DLLs
2. connect to the RIGOL generator
3. configure CH1 and CH2
4. set Shamrock grating and center wavelength
5. scan phase from start to stop
6. acquire one spectrum per repetition through the camera SDK
7. save one ASCII file per repetition

## Important Notes

- `trigger_mode=0` means internal trigger, `trigger_mode=1` means external trigger in the Andor SDK examples.
- The current implementation uses FVB mode to save one 1D spectrum per acquisition.
- Real hardware timing still needs one on-machine validation run.
