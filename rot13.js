window.addEventListener("DOMContentLoaded", main);
function main()
{
    // btRot13.addEventListener("click", rot13);
    iRot13.addEventListener("input", rot13);
}
function rot13()
{
    var rots="";
    var rot13=iRot13.value;
    for (var i=0; i<rot13.length;i++) {
        var ch =rot13[i];
        if (ch>="A" && ch <="Z")
             ch=String.fromCharCode((ch.charCodeAt(0)-"A".charCodeAt(0)+13) % 26 + "A".charCodeAt(0));
        if (ch>="a" && ch <="z") 
             ch=String.fromCharCode((ch.charCodeAt(0)-"a".charCodeAt(0)+13) % 26 + "a".charCodeAt(0));
        rots+=ch;
    }
    sRot13.innerHTML=rots;
}
