#ifndef __STRUCTURES_H__
#define __STRUCTURES_H__

// forces the structures to be 4byte aligned
#pragma pack(4)

typedef [uuid(01000000-7720-0045-7FFF-7ACDC6661234)]
struct MintDynData {
    long        cbsz;
    long        ArrPtr; //each array item: {long size;long valptr;}
    long        lpStrPtr_Name;
} MintDynData;
typedef [uuid(02000000-7720-0045-7FFF-7ACDC6661234)]
struct MintFileHeader {
    long        MintValidationKey; //Mint
    long        FileType; //ex: Cnfg
    long        Version; //ex: 1 means: 0.0.0.1
    long        StructureBegin_PTR; //NULL(able)
    /* Meta data here... */
    long        DataRecordLength;
    MintDynData RecordStructure;
} MintFileHeader;
typedef [uuid(03000000-7720-0045-7FFF-7ACDC6661234)]
struct MintFileHeaderP2 {
    long        cbsz;
    byte        CreationDate[8];
    byte        hshUsername[32];
} MintFileHeaderP2;

#endif //__STRUCTURES_H__