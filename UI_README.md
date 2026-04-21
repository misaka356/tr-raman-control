# UI Version

The UI version now provides:

- editable experiment parameters
- directory picker for the Andor SDK runtime root
- one `RIGOL VISA Resource` selector
- `Scan VISA` button to list available VISA resources
- `Start Generator` button for generator-only debugging
- `Stop Generator` button for generator-only stop
- `Debug Andor` button for spectrometer-only debugging
- `Run Experiment` for integrated scans
- save and load config
- live log window

The generator start logic is:

1. configure CH1 and CH2
2. enable triggered infinite burst on both channels
3. set both channels to waiting-trigger state with `:INITiate:IMMediate:ALL`
4. fire one shared `*TRG`

End-user workflow:

1. double-click `start_tr_raman.bat`
2. confirm `Andor SDK Root` points to `vendor\andor_sdk`
3. scan and select the correct VISA resource
4. use `Start Generator` or `Debug Andor` for separate debugging
5. click `Run Experiment` for the full automated scan
