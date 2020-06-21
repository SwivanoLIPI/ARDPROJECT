#define SensorPin 0          //pH meter Analog output to Arduino Analog Input 0
#include <Oxygen.h>
#include <MutichannelGasSensor.h>
#include <math.h>
#include <SoftwareSerial.h>
#include "rgb_lcd.h"
#include <SPI.h>
#include <OneWire.h> 
#include <DallasTemperature.h>
/********************************************************************/
// Data wire is plugged into pin 2 on the Arduino 
#define ONE_WIRE_BUS 2 
/********************************************************************/
// Setup a oneWire instance to communicate with any OneWire devices  
// (not just Maxim/Dallas temperature ICs) 
OneWire oneWire(ONE_WIRE_BUS); 
/********************************************************************/
// Pass our oneWire reference to Dallas Temperature. 
DallasTemperature sensors(&oneWire);
//#include <Wire.h>

rgb_lcd lcd;
const int colorR = 255;
const int colorG = 255;
const int colorB = 0;

unsigned long int avgValue;  //Store the average value of the sensor feedback
float b;
int buf[10],temp;

void setup()
{
  lcd.begin(16, 2);
  lcd.setRGB(colorR, colorG, colorB);
  pinMode(13,OUTPUT);  
  Serial.begin(9600);  
  sensors.begin(); 
  //Serial.println("Ready");    //Test the serial monitor
}
void loop()

{sensors.requestTemperatures(); 
delay(1000);
  for(int i=0;i<10;i++)       //Get 10 sample value from the sensor for smooth the value
  { 
    buf[i]=analogRead(SensorPin);
    delay(10);
  }
  for(int i=0;i<9;i++)        //sort the analog from small to large
  {
    for(int j=i+1;j<10;j++)
    {
      if(buf[i]>buf[j])
      {
        temp=buf[i];
        buf[i]=buf[j];
        buf[j]=temp;
      }
    }
  }
  avgValue=0;
  for(int i=2;i<8;i++)                      //take the average value of 6 center sample
    avgValue+=buf[i];
  float phValue=(float)avgValue*5.0/1024/6; //convert the analog into millivolt
  phValue=phValue+3.52;                      //convert the millivolt into pH value

  //Pengukuran Turbidity
  int sensorValue = analogRead(A1);// read the input on analog pin 0:
  float voltage = sensorValue * (5.0 / 1024.0); // Convert the analog reading (which goes from 0 - 1023) to a voltage (0 - 5V):
  
  // print out the value you read:
  lcd.setCursor(0, 0); 
  lcd.print("PH: "); 
  Serial.print(phValue,2);//ph
  Serial.print("|");
  Serial.print(voltage); //turbidity
  Serial.print("|");
  Serial.print(sensors.getTempCByIndex(0)-12.5);//temperature, nanti diperbaiki lagi
  Serial.println("|");
  lcd.print(phValue, 2);
  lcd.setCursor(10,0);
  if (phValue >=6.6 and phValue <=7.4 ){lcd.print("Normal");
  lcd.setCursor(0,1);
  lcd.print("Tidak Berbahaya");}
  else if (phValue <=6.6){lcd.print("Asam");}
  else if (phValue >=7.5){lcd.print("Basa");}
  //lcd.print("Sts:");
  //lcd.print("Asam");
  digitalWrite(13, HIGH);       
  digitalWrite(13, LOW); 
}
