- se c'è un modulo tipo Module=MDL_PCSerial; ..\..\S9380 PCSerial\MDL_PCSerial.bas --> lo consideriamo come classe esterna, non va rinominato nulla di quel modulo, in quel modulo etc...!


- se scrive AS VARIANT non lo vogliamo
  - Public Const RIC_AL_PLC_POLL As Variant = (&H1 Or AL_FAILURE_MASK)
	- 
- come mai scrive AS LONG negli OR?
  - Public Const RIC_AL_PDXI1_LOWVOLTAGE As Long = &H30 Or AL_ALARM_MASK 

- - se una procedura di un modulo non è private allora è public --> aggiungiamolo
- anche i tipi di modulo, se non sono private sono Public, aggiungiamolo
- se una variabile di modulo è DIM allora diventa Private


- controllare se le function hanno AS qualcosa, se manca segnalo sugli errori
- CALL da togliere ?
  - Call ExecuteReadStatus(m_PollRequest.PollCmd) --> ExecuteReadStatus m_PollRequest.PollCmdù
  - Call SetDataNotValid --> SetDataNotValid
- nei cicli for se c'è STEP 1, si può togliere