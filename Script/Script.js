function Export2Doc(element, filename = '') {

  var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
  var postHtml = "</body></html>";

  var html = preHtml + document.getElementById(element).innerHTML + postHtml;

  var blob = new Blob(['\ufeff', html], {
      type: 'application/msword'
  });

  var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);


  filename = filename ? filename + '.doc' : 'document.doc';


  var downloadLink = document.createElement("a");

  document.body.appendChild(downloadLink);

  if (navigator.msSaveOrOpenBlob) {
      navigator.msSaveOrOpenBlob(blob, filename);
  } else {

      downloadLink.href = url;
      downloadLink.download = filename;
      downloadLink.click();
  }

  document.body.removeChild(downloadLink);
  window.localStorage.clear();
}

function storeData(dataEvenimentului,foto,video,foto_video,standard,premiun,delux,alex,razvan,sebi,monea,mihai,dodo,cristi_duta,madalin,coco,online,recomandare,nume_sot,telefon_sot,nume_sotie,telefon_sotie,oras,local,nr_invitati,nr_contract,s_civila,x40,x30,x25,x20,x12,val_contract,m1,avans,m2,rest_plata,m3,extraoptiuni,alte_detalii,img){
  //Setare data eveniment

  var x = dataEvenimentului.split('-');
  var t  = x[1]+'.'+x[2]+'.'+x[0];
  localStorage.setItem('Data',t);

  //Setare Serviciu
  if(foto==true){
  
  localStorage.setItem('Serviciu','Foto-Video');
  }
  else if(video==true)
  {
    
    localStorage.setItem('Serviciu','Foto');
  }
  else if(foto_video==true)
  {
    localStorage.setItem('Serviciu','Video');
  }
  
    //Setare pachete

  if(standard==true){
  
    localStorage.setItem('Pachet','Standard');
    }
    else if(premiun==true)
    {
      
      localStorage.setItem('Pachet','Premium');
    }
    else if(delux==true)
    {
      localStorage.setItem('Pachet','Delux');
    }



    //Creare echia foto

    if(alex==true){
  
      localStorage.setItem('Alex','Alex');
    }
    if(razvan==true){
  
      localStorage.setItem('Razvan','Razvan');
    }
    if(sebi==true){
  
      localStorage.setItem('Sebi','Sebi');
    }
    if(monea==true){
  
      localStorage.setItem('Monea','Monea');
    }
    var echiapa_foto="";
    if(localStorage.getItem('Alex')!=null)
    {
      echiapa_foto= localStorage.getItem('Alex');
    }
    if(localStorage.getItem('Razvan')!=null)
    {
      echiapa_foto=echiapa_foto+" "+localStorage.getItem('Razvan');
    }
    if(localStorage.getItem('Sebi')!=null)
    {
      echiapa_foto= echiapa_foto+" "+localStorage.getItem('Sebi');
    }
    if(localStorage.getItem('Monea')!=null)
    {
      echiapa_foto= echiapa_foto+" "+localStorage.getItem('Monea');
    }
    
    localStorage.setItem('Echipa-Foto',echiapa_foto);


    //Creare echia Video 

    if(mihai==true){
  
      localStorage.setItem('Mihai','Mihai');
    }
    if(dodo==true){
  
      localStorage.setItem('Dodo','Dodo');
    }
    if(cristi_duta==true){
  
      localStorage.setItem('Cristi Duta','Cristi Duta');
    }
    if(madalin==true){
  
      localStorage.setItem('Madalin','Madalin');
    }
    if(coco==true){
  
      localStorage.setItem('Coco','Coco');
    }
    var echiapa_video="";
    if(localStorage.getItem('Mihai')!=null)
    {
      echiapa_video= localStorage.getItem('Mihai');
    }
    if(localStorage.getItem('Dodo')!=null)
    {
      echiapa_video=echiapa_video+" "+localStorage.getItem('Dodo');
    }
    if(localStorage.getItem('Cristi Duta')!=null)
    {
      echiapa_video= echiapa_video+" "+localStorage.getItem('Cristi Duta');
    }
    if(localStorage.getItem('Madalin')!=null)
    {
      echiapa_video= echiapa_video+" "+localStorage.getItem('Madalin');
    }
    if(localStorage.getItem('Coco')!=null)
    {
      echiapa_video= echiapa_video+" "+localStorage.getItem('Coco');
    }
    
    localStorage.setItem('Echipa-Video',echiapa_video);


    if(online==true){
  
      localStorage.setItem('Online','Online');
      }
      else if(recomandare==true)
      {
        
        localStorage.setItem('Recomandare','Recomandare');
      }
      var intrare="";
      if(localStorage.getItem('Online')!=null){
        intrare=localStorage.getItem('Online');
      }
      if(localStorage.getItem('Recomandare')!=null){
        intrare=localStorage.getItem('Recomandare');
      }
      intrare=localStorage.setItem('Intrare',intrare);

     // Adaugam Local storage datele de contact ale clientilor

      localStorage.setItem('Nume_sot',nume_sot);
      localStorage.setItem('Telefon_sot',telefon_sot);
      localStorage.setItem('Nume_sotie',nume_sotie);
      localStorage.setItem('Telefon_sotie',telefon_sotie);


      //Adaugam detaliile contractului

      localStorage.setItem('Oras',oras);
      localStorage.setItem('Local',local);
      localStorage.setItem('Nr_invitati',nr_invitati);
      localStorage.setItem('Nr_contract',nr_contract);
      localStorage.setItem('Situatie_civila',s_civila);

      //Materiale client

      
      localStorage.setItem('40X30',x40);
      localStorage.setItem('30X30',x30);
      localStorage.setItem('25X25',x25);
      localStorage.setItem('20X20',x20);
      localStorage.setItem('12X12',x12);

      if(localStorage.getItem('40X30')>0)
      {
        var t1=localStorage.getItem('40X30');
        var x40x30="-  Buc. Fotobook 40X30";
        x40x30 =[x40x30.slice(0, 2), t1, x40x30.slice(2)].join('');
        localStorage.setItem('40X30',x40x30);
      }


      if(localStorage.getItem('30X30')>0)
      {
        var t2=localStorage.getItem('30X30');
        var x30x30="-  Buc. Fotobook 30X30";
        x30x30 =[x30x30.slice(0, 2), t2, x30x30.slice(2)].join('');
        localStorage.setItem('30X30',x30x30);
      }


      if(localStorage.getItem('25X25')>0)
      {
        var t3=localStorage.getItem('25X25');
        var x25x25="-  Buc. Fotobook 25X25";
        x25x25 =[x25x25.slice(0, 2), t3, x25x25.slice(2)].join('');
        localStorage.setItem('25X25',x25x25);
      }

      if(localStorage.getItem('20X20')>0)
      {
        var t4=localStorage.getItem('20X20');
        var x20x20="-  Buc. Fotobook 20X20";
        x20x20 =[x20x20.slice(0, 2), t4, x20x20.slice(2)].join('');
        localStorage.setItem('20X20',x20x20);
      }

     


      if(localStorage.getItem('12X12',x12)>0)
      {
        var t5=localStorage.getItem('12X12',x12);
        var x12x12="-  Buc. Fotobook 12X12";
        x12x12 =[x12x12.slice(0, 2), t5, x12x12.slice(2)].join('');
        localStorage.setItem('12X12',x12x12);
      }

      //Valoare Contract si Plata acestuia

    
      var v_c =val_contract+" "+m1;
      localStorage.setItem('Valoare_contr',v_c);
      var v_a =avans+" "+m2;
      localStorage.setItem('Av',v_a);
      var v_r =rest_plata+" "+m3;
      localStorage.setItem('Rest_pl',v_r);

      // detalii si extaoptiuni

      var ex_fin=extraoptiuni.replaceAll('-', '<br> -');
      localStorage.setItem('Extra_o',ex_fin);
      var ex_det=alte_detalii.replaceAll('-', '<br> -');
      localStorage.setItem('Alte_D',ex_det);





    // Adaugare imagini contract fizic

    
    
  
  alert(b64);
}



