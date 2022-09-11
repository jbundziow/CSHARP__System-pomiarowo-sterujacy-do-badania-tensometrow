#define cewka_przekaznik A1
#define czuj_przem A0
#define silnik_DIR 3
#define silnik_STEP 4
#define silnik_SLEEP 5
#define silnik_RESET 6
#define MS1 7
#define MS2 8
#define MS3 9
#define silnik_ENABLE 10
#define LED_2V 11
#define LED_4V 12
#define ILOSC_POMIAROW 10


//ZMIENNE GLOBALNE
char dane;
int odczyt_analog;
float napiecie;

bool czy_zasilono_mostek = false;

void setup() {
  Serial.begin(9600);
  pinMode(13,OUTPUT);
  pinMode(cewka_przekaznik,OUTPUT);
  pinMode(silnik_DIR,OUTPUT);
  pinMode(silnik_STEP,OUTPUT);
  pinMode(silnik_SLEEP,OUTPUT);
  pinMode(silnik_RESET,OUTPUT);
  pinMode(silnik_ENABLE,OUTPUT);
  pinMode(MS1,OUTPUT);
  pinMode(MS2,OUTPUT);
  pinMode(MS3,OUTPUT);
  pinMode(LED_2V,OUTPUT);
  pinMode(LED_4V,OUTPUT);
  digitalWrite(LED_4V,HIGH);

  //USTAWIENIE ROZDZIELCZOSCI KROKU (USTALENIE EKSPERYMENTALNIE)
  digitalWrite(MS1,HIGH);
  digitalWrite(MS2,HIGH);
  digitalWrite(MS3,LOW);
}

void odczyt_czujnik_przemieszczenia ()
{
  for (int n=0; n<ILOSC_POMIAROW; n++)
  {
  odczyt_analog += analogRead(czuj_przem);
  }
  //napiecie = odczyt_analog * (5.00/1024.00)/ILOSC_POMIAROW; //STARE
  napiecie = odczyt_analog/ILOSC_POMIAROW;
  Serial.print(napiecie);
}




void loop() {
  if (Serial.available())
  {
  dane = Serial.read();


switch (dane)
{

  case '0': //zas OFF (czyli 4V)
  digitalWrite(cewka_przekaznik,LOW);
  digitalWrite(LED_2V, LOW);
  digitalWrite(LED_4V, HIGH);
  czy_zasilono_mostek = false;
  break;

  case '1': //zas 4V
  if (czy_zasilono_mostek == false)
  {
  digitalWrite(cewka_przekaznik,LOW);
  digitalWrite(LED_2V, LOW);
  digitalWrite(LED_4V, HIGH);
  czy_zasilono_mostek = true;
  }
  break;
  
  case '2': //zas 2V
  if (czy_zasilono_mostek == false)
  {
  digitalWrite(cewka_przekaznik,HIGH);
  digitalWrite(LED_2V, HIGH);
  digitalWrite(LED_4V, LOW);
  czy_zasilono_mostek = true;
  }
  break;
  
  case '3': //odczyt czuj. przem.
  odczyt_analog = 0;
  napiecie = 0;
  odczyt_czujnik_przemieszczenia();
  break;
  
  case '4': //silnik gora
  odczyt_czujnik_przemieszczenia();
  
  digitalWrite(silnik_ENABLE,LOW); //LOW = silnik chodzi
  digitalWrite(silnik_SLEEP, HIGH); //HIGH = silnik chodzi
  digitalWrite(silnik_RESET,HIGH); //HIGH = silnik chodzi
  digitalWrite(silnik_DIR,LOW); //LOW = prawo, HIGH=lewo
  digitalWrite(13, HIGH); //zapal diode 13 w arduino
  
  digitalWrite(silnik_STEP,HIGH);
  delay(5);
  digitalWrite(silnik_STEP,LOW);
  delay(5);
  break;

  case '5': //silnik dol
  odczyt_czujnik_przemieszczenia();

  digitalWrite(silnik_ENABLE,LOW); //LOW = silnik chodzi
  digitalWrite(silnik_SLEEP, HIGH); //HIGH = silnik chodzi
  digitalWrite(silnik_RESET,HIGH); //HIGH = silnik chodzi
  digitalWrite(silnik_DIR,HIGH); //LOW = prawo, HIGH=lewo
  digitalWrite(13, HIGH); //zapal diode 13 w arduino
  
  digitalWrite(silnik_STEP,HIGH);
  delay(5);
  digitalWrite(silnik_STEP,LOW);
  delay(5);
  break;

  case '6': //silniki OFF
  digitalWrite(silnik_ENABLE,HIGH); //LOW = silnik chodzi
  digitalWrite(silnik_SLEEP, LOW); //HIGH = silnik chodzi
  digitalWrite(silnik_RESET,LOW); //HIGH = silnik chodzi
  digitalWrite(silnik_DIR,LOW); //LOW = prawo, HIGH=lewo
  digitalWrite(13, LOW); //zgas diode 13 w arduino
  break;
}


}
else //wylacz silniki
{
  digitalWrite(silnik_ENABLE,HIGH); //LOW = silnik chodzi
  digitalWrite(silnik_SLEEP, LOW); //HIGH = silnik chodzi
  digitalWrite(silnik_RESET,LOW); //HIGH = silnik chodzi
  digitalWrite(silnik_DIR,LOW); //LOW = prawo, HIGH=lewo
  digitalWrite(13, LOW); //zgas diode 13 w arduino
}
digitalWrite(13, LOW); //zgas diode 13 w arduino
}
