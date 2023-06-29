import React, {useState} from "react";

const SupportedVersion = (props) => {
    const [SupportedVersion, setSupportedVersion] = useState(props.SupportedVersion);

    function isSetSupported(minVersion: string) {
        return Office.context.requirements.isSetSupported("Mailbox", minVersion)
    }

    return (
        <table id="supportedVersion"><tbody>
            <tr><td>Version</td><td>IsSupported</td></tr>
            <tr><td>1.1 </td><td>{`${isSetSupported("1.1")}`}</td></tr>
            <tr><td>1.2 </td><td>{`${isSetSupported("1.2")}`}</td></tr>
            <tr><td>1.3 </td><td>{`${isSetSupported("1.3")}`}</td></tr>
            <tr><td>1.4 </td><td>{`${isSetSupported("1.4")}`}</td></tr>
            <tr><td>1.5 </td><td>{`${isSetSupported("1.5")}`}</td></tr>
            <tr><td>1.6 </td><td>{`${isSetSupported("1.6")}`}</td></tr>
            <tr><td>1.7 </td><td>{`${isSetSupported("1.7")}`}</td></tr>
            <tr><td>1.8 </td><td>{`${isSetSupported("1.8")}`}</td></tr>
            <tr><td>1.9 </td><td>{`${isSetSupported("1.9")}`}</td></tr>
            <tr><td>1.10 </td><td>{`${isSetSupported("1.10")}`}</td></tr>
            <tr><td>1.11 </td><td>{`${isSetSupported("1.11")}`}</td></tr>
            <tr><td>1.12 </td><td>{`${isSetSupported("1.12")}`}</td></tr>
            <tr><td>1.13 </td><td>{`${isSetSupported("1.13")}`}</td></tr>
            <tr><td>1.14 </td><td>{`${isSetSupported("1.14")}`}</td></tr>
            <tr><td>1.15 </td><td>{`${isSetSupported("1.15")}`}</td></tr>
        </tbody></table>
    )
};

export default SupportedVersion;