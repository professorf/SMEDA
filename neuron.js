window.addEventListener("DOMContentLoaded", main); // wait UI Load
var network=[], training=[], lrate=0.1;
var istart, iend, hstart, hend, ostart, oend;

function main()
{
    // hardwire  neurons (ToDo: Connect-UI)
    var i1=new neuron();
    var i2=new neuron();
    var h1=new neuron();
    var h2=new neuron();
    var o=new neuron();

    ninp=2;
    nhid=2;
    nout=1;

    network=[i1,i2,h1,h2,o]; // in this sequence
                         // inputs-hidden-output

    istart=0;         iend=       ninp-1;
    hstart=ninp;      hend=hstart+nhid-1,
    ostart=ninp+nhid, oend=ostart+nout-1;
    // hardwire connections (ToDo: Connect-UI)
    i1.outputs=[h1,h2];
    i2.outputs=[h1,h2];
    h1.inputs=[i1, i2];
    h1.weights=[Math.random(),Math.random()];
    h1.outputs=[o];
    h2.inputs=[i1, i2];
    h2.weights=[Math.random(),Math.random()];
    h2.outputs=[o];
    o.inputs=[h1,h2];
    o.weights=[Math.random(),Math.random()];
    // test (ToDo: Training UI)
    /*
    i1.activation=1;
    i2.activation=1;
    h.calcNet();
    h.calcActivation();
    o.calcNet();
    o.calcActivation();
    */
    // setup training
    training=[[0,0,0],[0,1,1],[1,0,1],[1,1,0]];
    for (i=1; i<100000; i++)
        train();
    test();
}

test = function () {
    for (var r =0;r<training.length;r++)
    {
        // set inputs 
        for (var i=istart;i<=iend;i++) {
            network[i].activation=training[r][i];            
        }
        // propagate inputs to hidden
        for (var h=hstart;h<=hend;h++) {
            network[h].calcNet();
            network[h].calcActivation();
        }
        // propagate hidden to outputs
        for (var o=ostart;o<=oend;o++) {
            network[o].calcNet();
            network[o].calcActivation();            
        }   
        foo=0; // for a breakpoint
    }
}

train = function ()
{
    // load inputs

    for (var r =0;r<training.length;r++)
    {
        // set inputs 
        for (var i=istart;i<=iend;i++) {
            network[i].activation=training[r][i];            
        }
        // propagate inputs to hidden
        for (var h=hstart;h<=hend;h++) {
            network[h].calcNet();
            network[h].calcActivation();
        }
        // propagate hidden to outputs
        for (var o=ostart;o<=oend;o++) {
            network[o].calcNet();
            network[o].calcActivation();            
        }
        // now back propagate deltas: training to outputs
        for (var o=0;o<nout;o++) {
            var out=network[ostart+o];
            var netprime = out.activation*(1-out.activation);
            out.delta=(training[r][ninp+o]-out.activation) * netprime;
        }

        // now back propagate deltas: outputs to hidden
        for (var h=0;h<nhid;h++) {
            var hid = network[hstart+h];
            var netprime = hid.activation*(1-hid.activation);
            var sum=0;
            for (var ho=0;ho<hid.outputs.length;ho++) {
                var out=hid.outputs[ho];
                var del=out.delta;
                var ih=out.inputs.findIndex(element => element==hid);
                var wei=out.weights[ih];
                sum+=(del*wei);
            }
            hid.delta=netprime*sum;
        }

        // DO NOT back propagate deltas: hidden to outputs

        // now calculate the weight deltas, from hiddens to outputs 
        for (n=hstart;n<=oend;n++) {
            neu=network[n];
            for (i=0;i<neu.inputs.length;i++) {
                neu.wed[i]=lrate*neu.delta*neu.inputs[i].activation;
            }
        }

        // now update the weights based on the weds, from hiddens to outputs 
        for (n=hstart;n<=oend;n++) {
            neu=network[n];
            for (i=0;i<neu.inputs.length;i++) {
                neu.weights[i]+=neu.wed[i];
            }
        }
    }
}


neuron = function()
{
    this.inputs=[]; // input neurons
    this.weights=[]; // parallel w/inputs
    this.outputs=[]; // output neurons
    this.wed=[]; // parallel with weights
    this.net=0; // the neuron's internal state
    this.activation=0;   // the neuron's firing state 
    this.delta=0; // learning signal
    this.calcNet=function() 
    {
        this.net=0;
        for (var i=0;i<this.inputs.length;i++) {
            this.net+=(this.inputs[i].activation*this.weights[i]);
        }
    }    
    this.calcActivation = function () // logistic function
    {
        this.activation=1/(1+Math.exp(-this.net));
    }
}