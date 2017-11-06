/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');



var DynamicsWebApi = require('dynamics-web-api');

var AuthenticationContext = require('adal-node').AuthenticationContext;

var dynamicsWebApi = new DynamicsWebApi({ 
    webApiUrl: 'https://advancyaad.crm4.dynamics.com/api/data/v8.2/',
    onTokenRefresh: acquireToken
});
var authorityUrl = 'https://login.microsoftonline.com/94aeda88-8526-4ec8-b28f-fa67a055379f/oauth2/token';
var resource = 'https://advancyaad.crm4.dynamics.com';
var clientId = '1ae582b5-4b16-4b40-b180-0239e9b2b947';
var username = 'amokdad@advancyaad.onmicrosoft.com';
var password = 'p@ssw0rd2';
var adalContext = new AuthenticationContext(authorityUrl);

function acquireToken(dynamicsWebApiCallback){
    function adalCallback(error, token) {
        if (!error){
            dynamicsWebApiCallback(token);
        }
        else{
            
           // console.log(error);
        }
    }
    adalContext.acquireTokenWithUsernamePassword(resource, username, password, clientId, adalCallback);
}


// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
}); 
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    //appId: process.env.MicrosoftAppId,
    //appPassword: process.env.MicrosoftAppPassword,
    appId: "172a6955-ba7b-4e5e-90fe-2ac749701318",
    appPassword: "qjJVR5486/:ecgsmISMS8::",
    stateEndpoint: process.env.BotStateEndpoint,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());



var Recognizers = {
    arabicRecognizer : new builder.RegExpRecognizer( "Arabic", /العربية/i), 
    englishRecognizer : new builder.RegExpRecognizer( "English", /English/i)
}
var intents = new builder.IntentDialog({ recognizers: [    
    Recognizers.arabicRecognizer,
    Recognizers.englishRecognizer] 
,recognizeOrder:"series"})
.matches('English',(session, args) => {
    session.preferredLocale("en",function(err){
        if(!err){
            session.beginDialog("FirstDialog");
        }
     });
})

.matches('Arabic',(session, args) => {
    session.preferredLocale("ar",function(err){
        if(!err){
            session.beginDialog("FirstDialog");
        }
     });
})

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector,{
    localizerSettings: { 
        defaultLocale: "en" 
    }   
});

bot.dialog('/', intents);

bot.dialog("FirstDialog",[
    function(session){
       session.send("شكرا، سنقوم بالتواصل بلغتنا العربية.");
       builder.Prompts.choice(session, "يرجى الضغط على أي من الخيارات أدناه لكي أتمكن من مساعدتك أكثر." ,
       "أسئلة عامة|استشارة قانونية|استشارة تربوية|استشارة نفسية|استشارة اجتماعية|استشارة زوجية|لا أعرف، اريد مساعدة من أحد مستشاريكم",{listStyle: builder.ListStyle.button});

    },
    function(session,results){
        session.conversationData.q1 = results.response.entity;
        session.send("شكرا لاختيارك. أود أن نؤكد لك أننا سنسعى لتأمين الوفاق عند أسرتك بإذن الله.");
        builder.Prompts.text(session,'يرجى تزويدنا بالمزيد من التفاصيل عن المشكلة عبر طباعتها، أو بإمكانك الضغط على زر "تسجيل صوتي" لترك رسالة صوتية بسهولة.\n\nوسيقوم أحد مستشارينا بالتواصل معك بأسرع وقت ممكن');  
    },
    function(session,results){
       session.conversationData.q2 = session.message.text;
       session.beginDialog("getEmail");
    },
    function(session,results){
        session.conversationData.q3 = results.response;
        session.beginDialog("handeCRMEmail",{email:results.response}); 

    },
    function(session,results){
        var exist = results.response;
        if(exist){
            session.replaceDialog("UserExist");
        }else{
            session.replaceDialog("UserDoesntExist");
        }

    }
]);

function CreateContact(contact,crmCase){
    dynamicsWebApi.create(contact, "contacts").then(function (response) {
       var contactId = response;
       crmCase["customerid_contact@odata.bind"] = "https://advancyaad.crm4.dynamics.com/api/data/v8.2/contacts("+contactId+")";
       CreateCase(crmCase);

    })
    .catch(function (error){
        console.log(error);
    });
}
function CreateCase(crmCase){
    dynamicsWebApi.create(crmCase, "incidents").then(function (response) {
        //console.log('done');

    })
    .catch(function (error){
        console.log(error);
    });
}



bot.dialog("UserExist",[
    function(session){

        var crmCase = {
            title: session.conversationData.q1,
            description: session.conversationData.q2
        };
        if(session.conversationData.contactId == null)
            {
                var contact = {
                    firstname: session.conversationData.q5,
                    //lastname: session.conversationData.q5,
                    mobilephone: session.conversationData.q6,
                    emailaddress1: session.conversationData.q3
                };
                CreateContact(contact,crmCase);
            }
            else{
                crmCase["customerid_contact@odata.bind"] = "https://advancyaad.crm4.dynamics.com/api/data/v8.2/contacts("+session.conversationData.contactId+")";
                CreateCase(crmCase);
            }


        builder.Prompts.choice(session,'شكرا ' + session.conversationData.q5 + '، لقد قمنا بإرسال ملخص المشكلة ومعلومات إضافية الى بريدك الالكتروني أدناه، وبإمكانك أن تسألني في أي وقت عن حالة الشكوى إذا لم يصلك أي رد خلال يوم عمل واحد.\n\nلدينا مواد تعليمية وتوعية قد تودون قراءتها عن الحياة الزوجية والأسرية وكيفية التعامل معها.\n\nهل تود تصفح هذا المحتوى الخاص أو استلامها عبر البريد الالكتروني ؟','أريد أن أتصفح المحتوى الخاص|أود استلام المواد عبر البريد الالكتروني|الرجوع الى القائمة الرئيسية',{listStyle: builder.ListStyle.button});  
    },
    function(session,results){

    }
])
bot.dialog("UserDoesntExist",[
    function(session){
       builder.Prompts.choice(session, "يرجى اختيار الفئة العمرية" ,"أقل من 18 سنة|بين 18 و 24|بين 25 و 34|35 وما فوق",{listStyle: builder.ListStyle.button});
    },
    function(session,results){
        session.conversationData.q4 = results.response.entity;
        builder.Prompts.text(session,'يرجى كتابة اسمك الكامل');  
    },
    function(session,results){
        session.conversationData.q5 = session.message.text;
        session.beginDialog("getMobile");
    },
    function(session,results){
        session.conversationData.q6 = session.message.text;
        session.beginDialog("UserExist");
    }
]);

bot.dialog("getMobile",[
    function(session,args){
        if (args && args.reprompt) {
            builder.Prompts.text(session, "عفوا، يجب أن يكون رقم الجوال 8 أرقام على الأقل، يرجى المحاولة من جديد.");
        } else {
        builder.Prompts.text(session, "ما هو رقم جوالك؟");
        }
    },
    function(session,results)
    {
        var re = /[0-9]{8}/;
        if(re.test(results.response))
            session.endDialogWithResult(results);
        else
            session.replaceDialog('getMobile', { reprompt: true });
    }
]);
bot.dialog("handeCRMEmail",[
    function(session,args){
        var email = args.email;
        dynamicsWebApi.retrieveAll("contacts", ["emailaddress1","fullname"], "emailaddress1 eq '" + email + "'").then(function (response) {
            var records = response.value;
            //console.log(JSON.stringify(records));
            var exist = records != null && records.length >= 1;
            if(exist)
                {
                session.conversationData.contactId = records[0].contactId;
                session.conversationData.q5 =  records[0].fullname;
                
                }
            else
                session.conversationData.contactId = null;

            session.endDialogWithResult({response:exist});
        })
        .catch(function (error){
            console.log(error);
        });
    }
]);

bot.dialog("getEmail",[
    function(session,args){
        if (args && args.reprompt) {
            builder.Prompts.text(session, "عفوا، هذا البريد الالكتروني غير صحيح، يرجى المحاولة من جديد.");
        } else {
            builder.Prompts.text(session,'شكرا، لقد تم حفظ رسالتك وسنقوم بمساعدتك والتواصل معك سريعا.\n\nأود أخذ بعض معلومات التواصل منك لنتمكن من التواصل معك ومساعدتك أكثر.\n\nبإمكانك الاطمئنان بأننا نحفظ السرية التامة لك ولأسرتك وسنقوم فقط بالتواصل معك مباشرة.\n\nيرجى كتابة بريدك الالكتروني أدناه لنتمكن من التواصل معك');          
        }
    },
    function(session,results)
    {
        var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
        if(re.test(results.response))
            session.endDialogWithResult(results);
        else
            session.replaceDialog('getEmail', { reprompt: true });
    }
]);

bot.dialog("setLanguageWithPic",[
    function(session){
        
        var msg = new builder.Message(session);
        msg.attachmentLayout(builder.AttachmentLayout.carousel);
        var txt = "Hi, I’m your Personal Family Consultant and I’m here to answer your questions and help you easily find what you’re looking for. \n\n Please choose your preferred language \n\nالسلام عليكم، أنا مستشارك الأسري الشخصي وسأقوم بالإجابة على جميع أسئلتك ومساعدتك على إيجاد مبتغاك بسهولة تامة. \n\nيرجى إختيار لغة التواصل التي تفضلها";
        msg.attachments([
        new builder.HeroCard(session)
            .title("WIFAQ")
            .text(txt)
            //.images([builder.CardImage.create(session, "https://www.manateq.qa/Style%20Library/MTQ/Images/logo.png")])
            .buttons([
                builder.CardAction.imBack(session, "English", "English"),
                builder.CardAction.imBack(session, "العربية", "العربية"),
            ])
        ]);
        builder.Prompts.choice(session, msg, "العربية|English");
    },
    function(session,results){
        /*
        var contact = {
            firstname: "ahmad",
            lastname: "mokdad",
            mobilephone: "55979683",
            emailaddress1: "ahmad.mokdad@live.com"
        }
        var crmCase = {
            title: "dsadsa",
            description: "dadsa"
        }
        CreateContact(contact,crmCase);
        //session.send(contactId);
       //var locale = program.Helpers.GetLocal(results.response.index);
       */
       session.conversationData.lang = "ar";
       session.preferredLocale("ar",function(err){
           if(!err){
              session.replaceDialog("FirstDialog");    
           }
       });
       
    }
])

bot.on('conversationUpdate', function (activity) {  
    if (activity.membersAdded) {
        activity.membersAdded.forEach((identity) => {
            if (identity.id === activity.address.bot.id) {
                   bot.beginDialog(activity.address, 'setLanguageWithPic');
             }
         });
    }
 });