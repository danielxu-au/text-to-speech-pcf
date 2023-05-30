import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as SpeechSDK from 'microsoft-cognitiveservices-speech-sdk';

export class TextToSpeech implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    // state attributes
    private _context : ComponentFramework.Context<IInputs>;
    private _isInitialised : boolean = false;
    private _notifyOutputChanged: () => void;

    // property attributes
    private _text : string = "";
    private _state : string = "waiting";
    private _subscriptionKey : string;
    private _region : string;
    private _language : string;
    private _voice : string = "en-US-ChristopherNeural";
    private _autoSpeak : boolean = false;

    // ui attributes
    private _container : HTMLDivElement;
    private _buttonDiv : HTMLDivElement;

    /**
     * Empty constructor.
     */
    constructor()
    {
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        this._context = context;
        this._context.mode.trackContainerResize(true);
        this._container = container;

        // save the notifyOutputChanged
        this._notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        console.log("TextToSpeech:updateView() called");

        //has anything changed?  If not, bug out
        if(this._text === context.parameters.text.raw && 
           this._state === context.parameters.state.raw &&
           this._subscriptionKey === context.parameters.subscriptionKey.raw &&
           this._region === context.parameters.region.raw &&
           this._language === context.parameters.language.raw &&
           this._voice === context.parameters.voice.raw &&
           this._autoSpeak === context.parameters.autoSpeak.raw) {
            console.log("TextToSpeech:updateView() nothing has changed, ignoring the call");
            return;
        }

        // update the properties
        this._text = context.parameters.text.raw as string;
        this._state = context.parameters.state.raw as string;
        this._subscriptionKey = context.parameters.subscriptionKey.raw as string;
        this._region = context.parameters.region.raw as string;
        this._language = context.parameters.language.raw as string;
        this._voice = context.parameters.voice.raw as string;
        this._autoSpeak = context.parameters.autoSpeak.raw as boolean;

        // Add code to update control view
        if(!this._isInitialised) {        
            console.log(`adding icon to container`);
            // create the translation div & button
            this._buttonDiv = document.createElement("div");
            this._buttonDiv.id = `button-div`;
            this._buttonDiv.className = `button-div`;
            this._buttonDiv.style.width = `100%`;
            this._buttonDiv.style.height = `100%`;
            this._buttonDiv.style.cursor = `pointer`; 
            this._buttonDiv.innerHTML = `<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" viewBox="0 0 1024 1024" fill="none" xmlns="http://www.w3.org/2000/svg"> <path d="M571.5 269.085V766.017C571.5 792.832 540.496 807.755 519.538 791.027L419.257 710.99C413.588 706.465 406.549 704 399.295 704H252C181.308 704 124 646.692 124 576V480.5C124 409.808 181.308 352.5 252 352.5H397.274C405.744 352.5 413.868 349.142 419.867 343.163L516.908 246.423C537.081 226.312 571.5 240.6 571.5 269.085Z" fill="black"/> <path d="M683.5 326C743.007 374.595 781 448.541 781 531.361C781 614.181 743.007 688.127 683.5 736.722" stroke="black" stroke-width="48" stroke-linecap="round" stroke-linejoin="round"/> <path d="M624.5 435C654.406 459.255 673.5 496.163 673.5 537.5C673.5 578.837 654.406 615.745 624.5 640" stroke="black" stroke-width="48" stroke-linecap="round" stroke-linejoin="round"/> <path d="M781.5 281C854.129 340.158 900.5 430.178 900.5 531C900.5 631.822 854.129 721.842 781.5 781" stroke="black" stroke-width="48" stroke-linecap="round" stroke-linejoin="round"/> </svg>`;
            this._buttonDiv.addEventListener('click', this.startSpeakingREST.bind(this));
            this._container.appendChild(this._buttonDiv);  
            this._isInitialised = true;
        } else {
            this._buttonDiv.innerHTML = `<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" viewBox="0 0 1024 1024" fill="none" xmlns="http://www.w3.org/2000/svg"> <path d="M571.5 269.085V766.017C571.5 792.832 540.496 807.755 519.538 791.027L419.257 710.99C413.588 706.465 406.549 704 399.295 704H252C181.308 704 124 646.692 124 576V480.5C124 409.808 181.308 352.5 252 352.5H397.274C405.744 352.5 413.868 349.142 419.867 343.163L516.908 246.423C537.081 226.312 571.5 240.6 571.5 269.085Z" fill="black"/> <path d="M683.5 326C743.007 374.595 781 448.541 781 531.361C781 614.181 743.007 688.127 683.5 736.722" stroke="black" stroke-width="48" stroke-linecap="round" stroke-linejoin="round"/> <path d="M624.5 435C654.406 459.255 673.5 496.163 673.5 537.5C673.5 578.837 654.406 615.745 624.5 640" stroke="black" stroke-width="48" stroke-linecap="round" stroke-linejoin="round"/> <path d="M781.5 281C854.129 340.158 900.5 430.178 900.5 531C900.5 631.822 854.129 721.842 781.5 781" stroke="black" stroke-width="48" stroke-linecap="round" stroke-linejoin="round"/> </svg>`;
        }
        if(this._autoSpeak) this.startSpeakingREST();
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {
            "state" : this._state,
            "autoSpeak" : this._autoSpeak
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }


    public startSpeakingREST() : void
    {
        console.log("startSpeaking called");
        
        // check that we have text and are not already speaking
        if(this._text === "" || this._state === "speaking") {
            return;
        }

        // set state to speaking
        this._state = "speaking";
        this._autoSpeak = false;
        this._notifyOutputChanged();

        // create the async function
        const restAction = async () => {
            const response = await fetch(`https://${this._region}.tts.speech.microsoft.com/cognitiveservices/v1`, {
                method : "post",
                headers: {
                    "Ocp-Apim-Subscription-Key":`${this._subscriptionKey}`,
                    "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm",
                    "Content-Type": "application/ssml+xml"
                },
                body: `<speak version='1.0' xml:lang='${this._language}'><voice xml:lang='${this._language}' xml:gender='Male' name='${this._voice}'>${this._text}</voice></speak>`
            }).then(result => result.blob())
            .then(blob => {
                const audioURL = URL.createObjectURL(blob);
                const audio = new Audio(audioURL);
                audio.onended = () => {
                    this._state = "idle"; 
                    this._notifyOutputChanged();
                }
                audio.play();
            });
        }

        // do it
        restAction();
    }
}
